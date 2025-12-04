import 'dart:async';
import 'dart:io';
import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_selector/file_selector.dart';
import 'package:flutter/material.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'app_theme.dart';
import '../services/excel_file_service.dart';
import '../services/batch_processor.dart';
import '../services/excel_match_service.dart';

/// 文件解锁工具：简洁易用的解锁界面
class ExcelEditorPage extends StatefulWidget {
  const ExcelEditorPage({super.key});

  @override
  State<ExcelEditorPage> createState() => _ExcelEditorPageState();
}

class _ExcelEditorPageState extends State<ExcelEditorPage> {
  final _fileService = ExcelFileService();
  final _matchService = ExcelMatchService();

  // 模式切换
  bool _isMatchMode = false; // false = 解锁模式, true = 匹配模式

  // 匹配模式相关状态
  File? _draftFile; // 草稿文件
  File? _targetFile; // 目标文件
  bool _draftDragging = false;
  bool _targetDragging = false;
  bool _showInstructions = false; // 是否显示说明书（在 AppBar 中）
  bool _isAutoMatching = false; // 是否正在自动匹配
  bool _showOutputDirSettings = false; // 是否显示输出目录设置的详细内容

  bool _scheduleEnabled = false;
  DateTime? _scheduledDateTime; // 定时结束时间
  String? _scheduledInputDirPath; // 批量解锁源目录
  String? _scheduledOutputDirPath; // 批量解锁目标目录
  String? _outputDirExcel; // Excel 文件保存目录
  String? _outputDirPdf; // PDF 文件保存目录
  Timer? _scheduleTimer;
  int _batchExecutionCount = 0; // 批量解锁执行次数
  bool _dragging = false;
  bool _isLoading = false; // 用于单文件解锁和批量解锁
  bool _scheduleLoading = false; // 用于定时任务，与单文件解锁分离
  String? _statusMessage;
  String? _progressMessage;
  SharedPreferences? _prefs;
  bool _showScheduleSettings = false; // 控制定时功能区域的显示/隐藏（默认隐藏，用户点击右上角设置按钮才会显示）
  static const int _checkIntervalMinutes = 10; // 检查新文件的间隔时间（分钟）

  // 批量处理进度跟踪
  Timer? _progressUpdateTimer; // 用于节流UI更新
  DateTime? _lastProgressUpdate; // 上次更新UI的时间

  static const _prefsKeyInputDir = 'batch_input_dir';
  static const _prefsKeyOutputDir = 'batch_output_dir';
  static const _prefsKeyScheduledDate = 'scheduled_date_time';
  static const _prefsKeyOutputDirExcel = 'output_dir_excel'; // Excel 文件保存目录
  static const _prefsKeyOutputDirPdf = 'output_dir_pdf'; // PDF 文件保存目录
  static const _prefsKeyHasSetOutputDir = 'has_set_output_dir'; // 是否已设置过保存目录

  @override
  void initState() {
    super.initState();
    _loadBatchPreferences();
  }

  Future<void> _loadBatchPreferences() async {
    final prefs = await SharedPreferences.getInstance();
    _prefs = prefs;
    final input = prefs.getString(_prefsKeyInputDir);
    final output = prefs.getString(_prefsKeyOutputDir);
    final savedDateTimeStr = prefs.getString(_prefsKeyScheduledDate);
    final hasSetOutputDir = prefs.getBool(_prefsKeyHasSetOutputDir) ?? false;

    // 解析保存的定时时间
    final scheduledDateTime =
        savedDateTimeStr != null && savedDateTimeStr.isNotEmpty
            ? (() {
                try {
                  final dt = DateTime.parse(savedDateTimeStr);
                  // 检查时间是否已过期（不能是过去的时间）
                  if (dt.isBefore(DateTime.now())) {
                    prefs.remove(_prefsKeyScheduledDate);
                    return null;
                  }
                  return dt;
                } catch (e) {
                  debugPrint('[ExcelEditorPage] 解析保存的时间失败: $e');
                  return null;
                }
              })()
            : null;

    final excelDir = prefs.getString(_prefsKeyOutputDirExcel);
    final pdfDir = prefs.getString(_prefsKeyOutputDirPdf);

    setState(() {
      _scheduledInputDirPath = (input == null || input.isEmpty) ? null : input;
      _scheduledOutputDirPath =
          (output == null || output.isEmpty) ? null : output;
      _outputDirExcel =
          (excelDir == null || excelDir.isEmpty) ? null : excelDir;
      _outputDirPdf = (pdfDir == null || pdfDir.isEmpty) ? null : pdfDir;
      _scheduledDateTime = scheduledDateTime;
    });

    // 如果定时任务已启用且有有效时间，重新启动定时任务
    if (_scheduleEnabled && _scheduledDateTime != null) {
      _startSchedule();
    }

    // 首次打开应用时提示选择保存位置
    if (!hasSetOutputDir && mounted) {
      WidgetsBinding.instance.addPostFrameCallback((_) {
        _showFirstTimeSetupDialog();
      });
    }
  }

  Future<void> _persistBatchPreferences() async {
    final prefs = _prefs ?? await SharedPreferences.getInstance();
    await prefs.setString(_prefsKeyInputDir, _scheduledInputDirPath ?? '');
    await prefs.setString(_prefsKeyOutputDir, _scheduledOutputDirPath ?? '');
    await prefs.setString(_prefsKeyOutputDirExcel, _outputDirExcel ?? '');
    await prefs.setString(_prefsKeyOutputDirPdf, _outputDirPdf ?? '');
    if (_scheduledDateTime != null) {
      await prefs.setString(
          _prefsKeyScheduledDate, _scheduledDateTime!.toIso8601String());
    } else {
      await prefs.remove(_prefsKeyScheduledDate);
    }
  }

  /// 检查目录是否可写
  Future<bool> _canWriteToDirectory(String path) async {
    return (() async {
      final dir = Directory(path);
      if (!await dir.exists()) {
        await dir.create(recursive: true);
      }
      final testFile = File(
          '${dir.path}/.__perm_test_${DateTime.now().microsecondsSinceEpoch}');
      await testFile.writeAsString('test', flush: true);
      await testFile.delete();
      return true;
    })()
        .catchError((e) {
      debugPrint('[ExcelEditorPage] 无法写入目录 ($path)，可能缺少权限。错误: $e');
      return false;
    });
  }

  /// 请求存储权限（Android 平台）
  Future<bool> _requestStoragePermission() async {
    if (!Platform.isAndroid) return true;

    return (() async {
      if (await Permission.storage.isGranted) return true;

      final status = await Permission.storage.request();
      if (status.isGranted) return true;

      if (status.isPermanentlyDenied) {
        await openAppSettings();
      }
      return false;
    })()
        .catchError((e) {
      debugPrint('[ExcelEditorPage] _requestStoragePermission() ERROR: $e');
      return false;
    });
  }

  /// 解锁单个文件
  Future<void> _unlockSingleFile(File file) async {
    if (_isLoading) return;

    _updateLoadingState(
      isLoading: true,
      statusMessage: null,
      progressMessage: '正在解锁文件：${file.uri.pathSegments.last}...',
    );

    return _handleError(() async {
      final extension = file.path.toLowerCase().split('.').last;
      final fileTypeLabel = _getFileTypeLabel(extension);

      _updateLoadingState(progressMessage: '正在解锁 $fileTypeLabel...');

      // 解锁文件，处理特殊异常
      final unlocked = await _fileService.unlockFile(file).catchError(
        (error) {
          if (error is XlsFormatNotSupportedException) {
            _updateLoadingState(isLoading: false, progressMessage: null);
            _showXlsFormatDialog();
            throw error; // 重新抛出以终止流程
          }
          throw error;
        },
      );

      final finalExtension = unlocked.path.toLowerCase().split('.').last;
      final originalName = file.uri.pathSegments.last;
      final suggestedName = originalName.replaceFirst(
        RegExp(r'\.(xlsx|pdf)$', caseSensitive: false),
        '_unlocked.$finalExtension',
      );

      // 确定保存目录和文件路径
      final baseOutputDir = _getOutputDirForExtension(extension);
      final outFile = await _determineOutputFile(
        baseOutputDir: baseOutputDir,
        suggestedName: suggestedName,
        originalName: originalName,
        extension: finalExtension,
      );

      if (outFile == null) {
        _updateLoadingState(
          isLoading: false,
          statusMessage: '已解锁，但未选择保存位置',
          progressMessage: null,
        );
        return;
      }

      // 保存文件
      await _saveUnlockedFile(unlocked, outFile);

      _updateLoadingState(
        isLoading: false,
        statusMessage:
            '解锁成功：${outFile.uri.pathSegments.last}${_scheduledOutputDirPath != null ? '（已保存到批量解锁目标目录）' : ''}',
        progressMessage: null,
      );
    });
  }

  /// 获取文件类型标签
  String _getFileTypeLabel(String extension) {
    return switch (extension) {
      'xlsx' => 'Excel 文件',
      'pdf' => 'PDF 文件',
      _ => '文件',
    };
  }

  /// 根据扩展名获取对应的输出目录
  String? _getOutputDirForExtension(String extension) {
    return switch (extension) {
      'xlsx' => _outputDirExcel ?? _scheduledOutputDirPath,
      'pdf' => _outputDirPdf ?? _scheduledOutputDirPath,
      _ => _scheduledOutputDirPath,
    };
  }

  /// 确定输出文件路径（自动保存或用户选择）
  Future<File?> _determineOutputFile({
    String? baseOutputDir,
    required String suggestedName,
    required String originalName,
    required String extension,
  }) async {
    // 尝试自动保存
    if (baseOutputDir != null && baseOutputDir.isNotEmpty) {
      if (await _canWriteToDirectory(baseOutputDir)) {
        debugPrint('[ExcelEditorPage] 使用保存目录自动保存: $baseOutputDir');
        return _generateUniqueFilePath(
          baseDir: baseOutputDir,
          suggestedName: suggestedName,
          originalName: originalName,
          extension: extension,
        );
      }
      await _showInfoDialog(
        '没有写入权限',
        '应用无法写入下列目录：\n$baseOutputDir\n\n'
            '原因：macOS 仅对通过文件对话框选择的目录授予临时权限。\n\n'
            '请重新选择保存位置或在设置中重新指定保存目录。',
      );
    }

    // 用户选择保存位置
    final typeGroup = _getTypeGroupForExtension(extension);
    final location = await getSaveLocation(
      acceptedTypeGroups: [typeGroup],
      suggestedName: suggestedName,
    );

    return location?.path != null && location!.path.isNotEmpty
        ? File(location.path)
        : null;
  }

  /// 生成唯一的文件路径（处理文件已存在的情况）
  File _generateUniqueFilePath({
    required String baseDir,
    required String suggestedName,
    required String originalName,
    required String extension,
  }) {
    final targetPath =
        '$baseDir${baseDir.endsWith('/') ? '' : '/'}$suggestedName';
    final targetFile = File(targetPath);

    if (targetFile.existsSync()) {
      final ts = DateTime.now().millisecondsSinceEpoch;
      final uniqueName = originalName.replaceFirst(
        RegExp(r'\.(xlsx|pdf)$', caseSensitive: false),
        '_unlocked_$ts.$extension',
      );
      return File('$baseDir${baseDir.endsWith('/') ? '' : '/'}$uniqueName');
    }

    return targetFile;
  }

  /// 获取文件类型组
  XTypeGroup _getTypeGroupForExtension(String extension) {
    return switch (extension) {
      'xlsx' => const XTypeGroup(label: 'Excel', extensions: <String>['xlsx']),
      'pdf' => const XTypeGroup(label: 'PDF', extensions: <String>['pdf']),
      _ => const XTypeGroup(label: '文件', extensions: <String>[]),
    };
  }

  /// 保存解锁后的文件
  Future<void> _saveUnlockedFile(File unlocked, File targetFile) async {
    // 确保目录存在
    final directory = targetFile.parent;
    if (!await directory.exists()) {
      await directory.create(recursive: true);
    }

    // 写入文件
    final unlockedBytes = await unlocked.readAsBytes();
    await targetFile.writeAsBytes(unlockedBytes, flush: true);
  }

  /// 统一更新加载状态
  void _updateLoadingState({
    bool? isLoading,
    String? statusMessage,
    String? progressMessage,
  }) {
    if (!mounted) return;
    setState(() {
      if (isLoading != null) _isLoading = isLoading;
      if (statusMessage != null) _statusMessage = statusMessage;
      if (progressMessage != null) _progressMessage = progressMessage;
    });
  }

  /// 统一错误处理包装器
  Future<T> _handleError<T>(Future<T> Function() action) async {
    try {
      return await action();
    } catch (e, st) {
      debugPrint('[ExcelEditorPage] 操作失败: $e');
      debugPrint('[ExcelEditorPage] 堆栈: $st');
      if (mounted) {
        _updateLoadingState(
          isLoading: false,
          statusMessage: '操作失败：${e.toString()}',
          progressMessage: null,
        );
      }
      rethrow;
    }
  }

  /// 选择并解锁单个文件
  Future<void> _openAndUnlock() async {
    _updateLoadingState(
      statusMessage: null,
      progressMessage: '请选择要解锁的文件...',
    );

    return _handleError(() async {
      // Android 权限检查
      if (Platform.isAndroid) {
        final hasPermission = await _requestStoragePermission();
        if (!hasPermission) {
          _updateLoadingState(
            statusMessage: '需要存储权限才能选择文件',
            progressMessage: null,
          );
          return;
        }
      }

      // 选择文件
      final originalFile = await _fileService.pickFile();

      // 用户取消选择
      if (originalFile == null) {
        _updateLoadingState(
          statusMessage: '未选择文件',
          progressMessage: null,
        );
        return;
      }

      // 调用统一的解锁方法（会设置 loading 状态）
      await _unlockSingleFile(originalFile);
    });
  }

  /// 执行批量解锁（可重复执行）
  Future<void> _executeBatchUnlock() async {
    if (_scheduledInputDirPath == null || _scheduledOutputDirPath == null) {
      await _showInfoDialog('缺少目录', '请先选择批量解锁的源目录和目标目录。');
      return;
    }

    await _runBatchUnlockOnce(
      _scheduledInputDirPath!,
      _scheduledOutputDirPath!,
      fromManualTrigger: true,
    );
  }

  /// 批量解锁核心逻辑（优化版：使用并发处理提升性能）
  ///
  /// 采用并发处理机制，同时处理多个文件，但通过 BatchProcessor 控制并发数量
  /// 避免资源过度消耗和 UI 卡顿
  Future<void> _runBatchUnlockOnce(
    String inputDirPath,
    String outputDirPath, {
    bool fromSchedule = false,
    bool fromManualTrigger = false,
  }) async {
    // 增加执行计数
    setState(() {
      _batchExecutionCount++;
    });

    final inputDir = Directory(inputDirPath);
    final supportedExtensions = ['xlsx', 'pdf'];

    // 递归检索所有子目录中的文件
    debugPrint('[ExcelEditorPage] 开始递归扫描目录: $inputDirPath');
    final files = <File>[];

    // 递归遍历所有文件（扫描阶段）
    await for (final entity in inputDir.list(recursive: true)) {
      if (entity is File) {
        final extension = entity.path.toLowerCase().split('.').last;
        if (supportedExtensions.contains(extension)) {
          files.add(entity);
        }
      }
    }

    debugPrint('[ExcelEditorPage] 共找到 ${files.length} 个支持的文件（递归检索）');

    final triggerType = fromSchedule ? '定时' : (fromManualTrigger ? '手动' : '');

    if (files.isEmpty) {
      setState(() {
        _statusMessage = '$triggerType批量解锁：源目录中没有找到支持的文件';
      });
      return;
    }

    // 初始化进度跟踪
    _lastProgressUpdate = DateTime.now();

    setState(() {
      _isLoading = true;
      _progressMessage =
          '$triggerType批量解锁中...（共 ${files.length} 个文件，第 $_batchExecutionCount 次执行）';
    });

    // 创建批量处理器，限制并发数为 3（可调整）
    // 并发数设置为 3 可以在保证处理效率的同时，避免系统资源过度消耗
    final processor = BatchProcessor(maxConcurrency: 3);

    try {
      // 使用并发处理器批量处理文件
      // BatchProcessor 会控制同时处理的文件数量，避免资源耗尽
      final results = await processor.processBatch<File, File>(
        items: files,
        processor: (file, index) async {
          // 处理单个文件：解锁并保存
          return await _processBatchFile(
            file: file,
            outputDirPath: outputDirPath,
            index: index,
            total: files.length,
            fromSchedule: fromSchedule,
          );
        },
        onProgress: (completed, total, current) {
          // 进度更新回调（带节流机制，避免频繁更新 UI）
          _updateProgressThrottled(
            completed: completed,
            total: total,
            current: current,
            fromSchedule: fromSchedule,
            triggerType: triggerType,
          );
        },
      );

      // 处理完成后，统计成功和失败数量
      // BatchProcessor 返回成功处理的文件列表（失败的文件会被过滤掉）
      final successCount = results.length;
      final failCount = files.length - successCount;

      setState(() {
        _isLoading = false;
        _statusMessage =
            '$triggerType批量解锁完成（第 $_batchExecutionCount 次）：成功 $successCount，失败 $failCount';
        _progressMessage = null;
      });
    } catch (e, st) {
      debugPrint('[ExcelEditorPage] 批量解锁过程出错: $e');
      debugPrint('[ExcelEditorPage] 堆栈: $st');
      setState(() {
        _isLoading = false;
        _statusMessage = '$triggerType批量解锁失败：$e';
        _progressMessage = null;
      });
    } finally {
      // 清理进度更新定时器
      _progressUpdateTimer?.cancel();
      _progressUpdateTimer = null;
    }
  }

  /// 处理批量解锁中的单个文件
  ///
  /// 将文件解锁逻辑提取为独立方法，便于并发处理
  Future<File> _processBatchFile({
    required File file,
    required String outputDirPath,
    required int index,
    required int total,
    required bool fromSchedule,
  }) async {
    // 解锁文件
    final unlocked = await _fileService.unlockFile(file);
    final originalName = file.uri.pathSegments.last;
    final finalExtension = unlocked.path.toLowerCase().split('.').last;

    // 生成目标文件路径（处理文件已存在的情况）
    final normalizedDir =
        outputDirPath.endsWith('/') ? outputDirPath : '$outputDirPath/';
    final targetFile = await _generateBatchOutputFile(
      baseDir: normalizedDir,
      originalName: originalName,
      extension: finalExtension,
    );

    // 确保目录存在并写入文件
    await targetFile.parent.create(recursive: true);
    final unlockedBytes = await unlocked.readAsBytes();
    await targetFile.writeAsBytes(unlockedBytes, flush: true);

    return targetFile;
  }

  /// 生成批量处理的输出文件路径（处理文件已存在的情况）
  Future<File> _generateBatchOutputFile({
    required String baseDir,
    required String originalName,
    required String extension,
  }) async {
    var targetName = originalName.replaceFirst(
      RegExp(r'\.(xlsx|pdf)$', caseSensitive: false),
      '_unlocked.$extension',
    );
    var targetFile = File('$baseDir$targetName');

    // 如果文件已存在，添加时间戳
    if (await targetFile.exists()) {
      final ts = DateTime.now().millisecondsSinceEpoch;
      targetName = originalName.replaceFirst(
        RegExp(r'\.(xlsx|pdf)$', caseSensitive: false),
        '_unlocked_$ts.$extension',
      );
      targetFile = File('$baseDir$targetName');
    }

    return targetFile;
  }

  /// 节流更新进度信息（避免频繁触发 setState）
  ///
  /// 每 500ms 或每处理 3 个文件更新一次 UI，减少重绘频率
  void _updateProgressThrottled({
    required int completed,
    required int total,
    File? current,
    required bool fromSchedule,
    required String triggerType,
  }) {
    final now = DateTime.now();
    final shouldUpdate = _lastProgressUpdate == null ||
        now.difference(_lastProgressUpdate!) >=
            const Duration(milliseconds: 500) ||
        completed % 3 == 0 || // 每处理 3 个文件更新一次
        completed == total; // 最后一个文件必须更新

    if (shouldUpdate && mounted) {
      _lastProgressUpdate = now;
      final currentFileName = current?.uri.pathSegments.last ?? '';
      setState(() {
        _progressMessage =
            '${fromSchedule ? '[定时任务] ' : ''}正在处理 ($completed/$total)${currentFileName.isNotEmpty ? '：$currentFileName' : ''}';
      });
    }
  }

  void _startSchedule() {
    if (_scheduledInputDirPath == null || _scheduledOutputDirPath == null) {
      setState(() {
        _scheduleEnabled = false;
        _statusMessage = '请先设置源目录和目标目录';
      });
      return;
    }

    if (_scheduledDateTime == null) {
      setState(() {
        _scheduleEnabled = false;
        _statusMessage = '请先设置定时结束时间';
      });
      return;
    }

    // 检查时间是否是过去的时间
    final now = DateTime.now();
    if (_scheduledDateTime!.isBefore(now)) {
      setState(() {
        _scheduleEnabled = false;
        _statusMessage = '定时结束时间不能是过去的时间，请重新选择';
        _scheduledDateTime = null;
      });
      _persistBatchPreferences();
      return;
    }

    _scheduleTimer?.cancel();

    // 立即执行一次检查（检查是否有新文件）
    Future.microtask(() async {
      if (!_scheduleLoading && _scheduleEnabled && mounted) {
        final currentTime = DateTime.now();
        // 检查是否还在定时时间范围内
        if (currentTime.isBefore(_scheduledDateTime!)) {
          setState(() => _scheduleLoading = true);
          await _runBatchUnlockOnce(
            _scheduledInputDirPath!,
            _scheduledOutputDirPath!,
            fromSchedule: true,
          );
          if (mounted) {
            setState(() => _scheduleLoading = false);
          }
        }
      }
    });

    // 设置定期检查定时器（每分钟检查一次）
    _scheduleTimer = Timer.periodic(
      const Duration(minutes: _checkIntervalMinutes),
      (timer) async {
        final currentTime = DateTime.now();

        // 检查是否已超过结束时间
        if (currentTime.isAfter(_scheduledDateTime!) ||
            currentTime.isAtSameMomentAs(_scheduledDateTime!)) {
          // 到达结束时间，停止定时任务
          timer.cancel();
          setState(() {
            _scheduleEnabled = false;
            _statusMessage = '定时任务已结束（已到达结束时间）';
          });
          _persistBatchPreferences();
          return;
        }

        // 如果不在加载中且定时任务仍启用，检查新文件并执行解锁
        if (!_scheduleLoading && _scheduleEnabled && mounted) {
          setState(() => _scheduleLoading = true);
          await _runBatchUnlockOnce(
            _scheduledInputDirPath!,
            _scheduledOutputDirPath!,
            fromSchedule: true,
          );
          if (mounted) {
            setState(() => _scheduleLoading = false);
          }
        }
      },
    );

    // 格式化显示时间
    final dateStr =
        '${_scheduledDateTime!.year}-${_scheduledDateTime!.month.toString().padLeft(2, '0')}-${_scheduledDateTime!.day.toString().padLeft(2, '0')}';
    final timeStr =
        '${_scheduledDateTime!.hour.toString().padLeft(2, '0')}:${_scheduledDateTime!.minute.toString().padLeft(2, '0')}';
    setState(() {
      _statusMessage =
          '定时自动解锁已开启，将在 $dateStr $timeStr 结束（每 $_checkIntervalMinutes 分钟检查一次新文件）';
    });
  }

  void _stopSchedule() {
    _scheduleTimer?.cancel();
    _scheduleTimer = null;
    if (mounted) {
      setState(() {
        _scheduleEnabled = false; // 确保状态正确更新
        _scheduleLoading = false;
        _statusMessage = '⏸️ 定时任务已停止';
      });
      _persistBatchPreferences();
    }
  }

  Future<void> _selectScheduledInputDir() async {
    final String? path = await getDirectoryPath();
    if (path != null && path.isNotEmpty) {
      await _setInputDirectory(path);
    }
  }

  Future<void> _selectScheduledOutputDir() async {
    final String? path = await getDirectoryPath();
    if (path != null && path.isNotEmpty) {
      await _setOutputDirectory(path);
    }
  }

  /// 选择定时结束的时间
  Future<void> _selectScheduledDateTime() async {
    final now = DateTime.now();
    final firstDate = now; // 最早可以选择今天（不能选择过去的时间）
    // 移除10天限制，只限制不能选择过去的时间
    final lastDate = DateTime(now.year + 100); // 设置为很远未来的日期（约100年后）

    // 选择日期
    final selectedDate = await showDatePicker(
      context: context,
      initialDate: _scheduledDateTime ?? now,
      firstDate: firstDate,
      lastDate: lastDate,
      helpText: '选择定时结束日期（不能选择过去的时间）',
      builder: (context, child) {
        return Theme(
          data: Theme.of(context).copyWith(
            colorScheme: const ColorScheme.light(
              primary: brandPink,
              onPrimary: Colors.white,
              surface: Colors.white,
              onSurface: brandNavy,
            ),
          ),
          child: child!,
        );
      },
    );

    if (selectedDate == null) return;
    if (!mounted) return;

    // 选择时间
    final selectedTime = await showTimePicker(
      context: context,
      initialTime: _scheduledDateTime != null
          ? TimeOfDay(
              hour: _scheduledDateTime!.hour,
              minute: _scheduledDateTime!.minute,
            )
          : TimeOfDay.now(),
      helpText: '选择定时结束时间',
      builder: (context, child) {
        return Theme(
          data: Theme.of(context).copyWith(
            colorScheme: const ColorScheme.light(
              primary: brandPink,
              onPrimary: Colors.white,
              surface: Colors.white,
              onSurface: brandNavy,
            ),
          ),
          child: child!,
        );
      },
    );

    if (selectedTime == null) return;
    if (!mounted) return;

    // 组合日期和时间
    final scheduledDateTime = DateTime(
      selectedDate.year,
      selectedDate.month,
      selectedDate.day,
      selectedTime.hour,
      selectedTime.minute,
    );

    // 验证：时间不能是过去的时间
    if (scheduledDateTime.isBefore(now)) {
      if (!mounted) return;
      await _showInfoDialog('时间无效', '不能选择过去的时间作为结束时间，请选择今天或未来的时间。');
      return;
    }

    setState(() {
      _scheduledDateTime = scheduledDateTime;
    });
    await _persistBatchPreferences();

    // 如果定时任务已启用，重新启动
    if (_scheduleEnabled) {
      _stopSchedule();
      _startSchedule();
    }
  }

  Future<void> _setInputDirectory(String path) async {
    setState(() {
      _scheduledInputDirPath = path;
      _statusMessage = '已设置源目录：$path';
    });
    await _persistBatchPreferences();
    if (_scheduleEnabled) _startSchedule();
  }

  Future<void> _setOutputDirectory(String path) async {
    setState(() {
      _scheduledOutputDirPath = path;
      _statusMessage = '已设置目标目录：$path';
    });
    await _persistBatchPreferences();
    if (_scheduleEnabled) _startSchedule();
  }

  Future<void> _handleScheduleToggle(bool value) async {
    if (value) {
      if (_scheduledInputDirPath == null || _scheduledOutputDirPath == null) {
        await _showInfoDialog('缺少目录', '请先选择批量解锁的源目录和目标目录。');
        setState(() {});
        return;
      }
      if (_scheduledDateTime == null) {
        await _showInfoDialog('缺少时间', '请先选择定时结束时间。');
        setState(() {});
        return;
      }
      // 检查时间是否是过去的时间
      if (_scheduledDateTime!.isBefore(DateTime.now())) {
        await _showInfoDialog('时间无效', '定时结束时间不能是过去的时间，请重新选择。');
        setState(() {
          _scheduledDateTime = null;
        });
        await _persistBatchPreferences();
        return;
      }
      setState(() {
        _scheduleEnabled = true;
      });
      _startSchedule();
    } else {
      setState(() {
        _scheduleEnabled = false;
      });
      _stopSchedule();
    }
  }

  Future<void> _showInfoDialog(String title, String message) async {
    if (!mounted) return;
    await showDialog<void>(
      context: context,
      builder: (context) => AlertDialog(
        title: Row(
          children: [
            const Icon(Icons.info_outline),
            const SizedBox(width: 8),
            Text(title),
          ],
        ),
        content: Text(message),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('好的'),
          ),
        ],
      ),
    );
  }

  /// 首次打开应用时显示设置对话框
  Future<void> _showFirstTimeSetupDialog() async {
    if (!mounted) return;
    await showDialog<void>(
      context: context,
      barrierDismissible: false,
      builder: (context) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.settings, color: brandPink),
            SizedBox(width: 8),
            Text('欢迎使用文件解锁工具'),
          ],
        ),
        content: const Text(
          '为了更好地管理解锁后的文件，需要设置保存位置。\n\n'
          '可以选择：\n'
          '1. 设置统一的保存目录（所有文件保存在同一目录）\n'
          '2. 分别设置 Excel 和 PDF 的保存目录（文件分类保存）\n\n'
          '选择跳过，后续可以在设置中配置。\n\n'
          'qwq！！！！！！！哒哒、哒哒ovo',
        ),
        actions: [
          TextButton(
            onPressed: () {
              Navigator.of(context).pop();
              // 标记已设置过（即使选择跳过）
              _prefs?.setBool(_prefsKeyHasSetOutputDir, true);
            },
            child: const Text('稍后设置'),
          ),
          TextButton(
            onPressed: () {
              Navigator.of(context).pop();
              _showSelectOutputDirDialog();
            },
            style: TextButton.styleFrom(
              foregroundColor: brandPink,
            ),
            child: const Text('现在设置'),
          ),
        ],
      ),
    );
  }

  /// 显示选择保存目录的对话框
  Future<void> _showSelectOutputDirDialog() async {
    if (!mounted) return;
    await showDialog<void>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.folder, color: brandPink),
            SizedBox(width: 8),
            Text('选择保存位置'),
          ],
        ),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Text(
                '您可以选择：',
                style: TextStyle(fontWeight: FontWeight.bold),
              ),
              const SizedBox(height: 12),
              ElevatedButton.icon(
                onPressed: () async {
                  Navigator.of(context).pop();
                  await _selectUnifiedOutputDir();
                },
                icon: const Icon(Icons.folder_open),
                label: const Text('统一目录（所有文件保存到同一目录）'),
                style: ElevatedButton.styleFrom(
                  backgroundColor: brandNavy,
                  foregroundColor: Colors.white,
                  padding:
                      const EdgeInsets.symmetric(vertical: 12, horizontal: 16),
                ),
              ),
              const SizedBox(height: 12),
              ElevatedButton.icon(
                onPressed: () async {
                  Navigator.of(context).pop();
                  await _selectSeparateOutputDirs();
                },
                icon: const Icon(Icons.folder_copy),
                label: const Text('分类目录（Excel 和 PDF 分别保存）'),
                style: ElevatedButton.styleFrom(
                  backgroundColor: brandPink,
                  foregroundColor: Colors.white,
                  padding:
                      const EdgeInsets.symmetric(vertical: 12, horizontal: 16),
                ),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('取消'),
          ),
        ],
      ),
    );
  }

  /// 选择统一的保存目录
  Future<void> _selectUnifiedOutputDir() async {
    final String? path = await getDirectoryPath();
    if (path != null && path.isNotEmpty) {
      setState(() {
        _scheduledOutputDirPath = path;
        // 统一目录时，也设置Excel和PDF目录为同一目录
        _outputDirExcel = path;
        _outputDirPdf = path;
      });
      await _prefs?.setBool(_prefsKeyHasSetOutputDir, true);
      await _persistBatchPreferences();
      if (mounted) {
        _showInfoDialog('设置成功', '已设置统一的保存目录：$path\n\nExcel 和 PDF 文件都将保存到此目录。');
      }
    }
  }

  /// 分别选择 Excel 和 PDF 的保存目录
  Future<void> _selectSeparateOutputDirs() async {
    // 选择 Excel 保存目录
    final excelPath = await getDirectoryPath();
    if (excelPath == null || excelPath.isEmpty) return;

    // 选择 PDF 保存目录
    final pdfPath = await getDirectoryPath();
    if (pdfPath == null || pdfPath.isEmpty) return;

    setState(() {
      _outputDirExcel = excelPath;
      _outputDirPdf = pdfPath;
    });
    await _prefs?.setBool(_prefsKeyHasSetOutputDir, true);
    await _persistBatchPreferences();
    if (mounted) {
      _showInfoDialog(
        '设置成功',
        '已分别设置保存目录：\n\n'
            'Excel 文件目录：$excelPath\n'
            'PDF 文件目录：$pdfPath',
      );
    }
  }

  /// 显示 .xls 格式不支持提示弹窗
  Future<void> _showXlsFormatDialog() async {
    if (!mounted) return;
    await showDialog<void>(
      context: context,
      builder: (context) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.warning_amber_rounded, color: Colors.orange),
            SizedBox(width: 8),
            Text('不支持 .xls 格式'),
          ],
        ),
        content: const Text(
          '该应用不支持 .xls 格式文件。\n\n'
          '建议：尝试用 Microsoft Excel 打开该文件并另存为标准的 .xlsx 格式。',
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(context).pop(),
            child: const Text('知道了'),
          ),
        ],
      ),
    );
  }

  @override
  void dispose() {
    _scheduleTimer?.cancel();
    _progressUpdateTimer?.cancel();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context);
    final isErrorStatus = _statusMessage?.startsWith('解锁失败') == true;
    final isSuccessStatus = _statusMessage?.startsWith('解锁成功') == true;
    return Scaffold(
      backgroundColor: backgroundWhite,
      appBar: AppBar(
        title: Text(_isMatchMode ? 'Excel 匹配工具' : '文件解锁工具'),
        backgroundColor: brandNavy,
        foregroundColor: Colors.white,
        elevation: 0,
        actions: [
          // 使用说明按钮（仅在匹配模式显示）
          if (_isMatchMode)
            IconButton(
              icon: Icon(
                _showInstructions ? Icons.info : Icons.info_outline,
                color: Colors.white,
              ),
              onPressed: () {
                setState(() {
                  _showInstructions = !_showInstructions;
                });
                if (_showInstructions) {
                  _showInfoDialog(
                    '使用说明',
                    '1. 将草稿 Excel 文件拖到左侧区域（支持多 Sheet）\n'
                        '2. 将目标 Excel 文件拖到右侧区域（单 Sheet）\n'
                        '3. 设置输出目录\n'
                        '4. 拖拽文件后会自动执行匹配\n'
                        '5. 程序会自动识别列结构，根据 Item 和 Description 匹配\n'
                        '6. 匹配结果将保存到输出目录',
                  );
                }
              },
              tooltip: '使用说明',
            ),
          // 模式切换按钮
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 8.0),
            child: Row(
              children: [
                TextButton.icon(
                  onPressed: () {
                    setState(() {
                      _isMatchMode = !_isMatchMode;
                      // 切换模式时清空匹配模式的文件
                      if (_isMatchMode) {
                        _draftFile = null;
                        _targetFile = null;
                      }
                    });
                  },
                  icon: Icon(
                    _isMatchMode ? Icons.lock_open : Icons.compare_arrows,
                    color: Colors.white,
                    size: 18,
                  ),
                  label: Text(
                    _isMatchMode ? '解锁模式' : '匹配模式',
                    style: const TextStyle(color: Colors.white, fontSize: 14),
                  ),
                  style: TextButton.styleFrom(
                    backgroundColor: _isMatchMode
                        ? brandPink.withOpacity(0.3)
                        : brandPink.withOpacity(0.3),
                    padding:
                        const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                  ),
                ),
                const SizedBox(width: 8),
              ],
            ),
          ),
          IconButton(
            icon: Icon(
              _showScheduleSettings ? Icons.settings : Icons.settings_outlined,
              color: Colors.white,
            ),
            onPressed: () {
              setState(() {
                _showScheduleSettings = !_showScheduleSettings;
              });
            },
            tooltip: _showScheduleSettings ? '隐藏定时设置' : '显示定时设置',
          ),
        ],
      ),
      body: Column(
        children: [
          // 状态消息区域
          if (_statusMessage != null || _progressMessage != null)
            Container(
              width: double.infinity,
              padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
              color: isErrorStatus
                  ? brandRed.withOpacity(0.1)
                  : isSuccessStatus
                      ? successGreen.withOpacity(0.1)
                      : brandNavy.withOpacity(0.05),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  if (_statusMessage != null)
                    Row(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Icon(
                          isErrorStatus
                              ? Icons.error_outline
                              : isSuccessStatus
                                  ? Icons.check_circle_outline
                                  : Icons.info_outline,
                          color: isErrorStatus
                              ? brandRed
                              : isSuccessStatus
                                  ? successGreen
                                  : infoBlue,
                        ),
                        const SizedBox(width: 8),
                        Expanded(
                          child: Text(
                            _statusMessage!,
                            style: theme.textTheme.bodyMedium?.copyWith(
                              color: isErrorStatus
                                  ? brandRed
                                  : isSuccessStatus
                                      ? successGreen
                                      : theme.textTheme.bodyMedium?.color,
                              fontWeight: FontWeight.w600,
                            ),
                          ),
                        ),
                      ],
                    ),
                  if (_progressMessage != null) ...[
                    const SizedBox(height: 4),
                    Row(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        const Icon(Icons.autorenew, color: infoBlue),
                        const SizedBox(width: 8),
                        Expanded(
                          child: Text(
                            _progressMessage!,
                            style: theme.textTheme.bodySmall?.copyWith(
                              color: infoBlue,
                            ),
                          ),
                        ),
                      ],
                    ),
                  ],
                ],
              ),
            ),

          // 主要内容区域
          Expanded(
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: _isMatchMode
                    ? _buildMatchModeContent(theme)
                    : _buildUnlockModeContent(theme),
              ),
            ),
          ),
        ],
      ),
    );
  }

  /// 构建解锁模式的内容
  List<Widget> _buildUnlockModeContent(ThemeData theme) {
    return [
      // 大拖拽区域
      DropTarget(
        onDragEntered: (_) => setState(() => _dragging = true),
        onDragExited: (_) => setState(() => _dragging = false),
        onDragDone: (details) async {
          setState(() => _dragging = false);
          if (details.files.isEmpty) return;
          final path = details.files.first.path;

          if (await File(path).exists()) {
            final ext = path.toLowerCase().split('.').last;
            if (['xlsx', 'pdf'].contains(ext)) {
              await _unlockSingleFile(File(path));
            } else if (ext == 'xls') {
              await _showXlsFormatDialog();
            } else {
              setState(() {
                _statusMessage = '不支持的文件格式：.$ext（仅支持 .xlsx, .pdf）';
              });
            }
          } else if (await Directory(path).exists()) {
            await _setInputDirectory(path);
          }
        },
        child: Container(
          height: 320, // 从 200 增加到 280，使拖拽区域更大
          decoration: BoxDecoration(
            color: _dragging
                ? successGreen.withOpacity(0.15)
                : brandCream.withOpacity(0.3),
            borderRadius: BorderRadius.circular(16),
            border: Border.all(
              color: _dragging ? successGreen : brandNavy.withOpacity(0.3),
              width: _dragging ? 3 : 2,
            ),
          ),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Icon(
                _dragging ? Icons.cloud_download : Icons.cloud_upload,
                size: 80, // 从 64 增加到 80，使图标更大
                color: _dragging ? successGreen : brandNavy,
              ),
              const SizedBox(height: 20), // 从 16 增加到 20
              Text(
                _dragging ? '释放文件以解锁' : '拖拽文件到这里解锁',
                style: theme.textTheme.titleLarge?.copyWith(
                  color: brandNavy,
                  fontWeight: FontWeight.bold,
                  fontSize: 22, // 增加字体大小
                ),
              ),
              const SizedBox(height: 12), // 从 8 增加到 12
              Text(
                '支持 .xlsx、.pdf 格式',
                style: theme.textTheme.bodyMedium?.copyWith(
                  color: Colors.grey[600],
                  fontSize: 16, // 增加字体大小
                ),
              ),
            ],
          ),
        ),
      ),

      const SizedBox(height: 24),

      // 操作按钮
      Row(
        children: [
          Expanded(
            child: ElevatedButton.icon(
              onPressed: _isLoading ? null : _openAndUnlock,
              icon: const Icon(
                Icons.folder_open,
                size: 28, // 增加图标大小
              ),
              label: const Text(
                '选择文件解锁',
                style: TextStyle(
                  fontSize: 18, // 增加字体大小
                  fontWeight: FontWeight.w600, // 增加字体粗细
                ),
              ),
              style: ElevatedButton.styleFrom(
                backgroundColor: brandNavy,
                foregroundColor: Colors.white,
                padding: const EdgeInsets.symmetric(
                  vertical: 20, // 从 16 增加到 20
                  horizontal: 24, // 增加水平内边距
                ),
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(12),
                ),
                elevation: 2, // 增加阴影效果
              ),
            ),
          ),
        ],
      ),

      const SizedBox(height: 32),

      // 批量解锁设置（仅在 _showScheduleSettings 为 true 时显示）
      if (_showScheduleSettings) ...[
        Card(
          elevation: 2,
          shape: RoundedRectangleBorder(
            borderRadius: BorderRadius.circular(12),
          ),
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    const Icon(Icons.storage, color: brandNavy),
                    const SizedBox(width: 8),
                    Text(
                      '批量解锁',
                      style: theme.textTheme.titleMedium?.copyWith(
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 12),
                Text(
                  '1. 选择源目录与目标目录\n2. 点击按钮立即批量解锁或开启定时任务',
                  style: theme.textTheme.bodySmall,
                ),
                const SizedBox(height: 16),
                Row(
                  children: [
                    Expanded(
                      child: OutlinedButton.icon(
                        onPressed: _isLoading ? null : _selectScheduledInputDir,
                        icon: const Icon(Icons.folder),
                        label: Text(
                          _scheduledInputDirPath != null
                              ? '源目录：${_scheduledInputDirPath!.split('/').last}'
                              : '选择源目录',
                          overflow: TextOverflow.ellipsis,
                        ),
                        style: OutlinedButton.styleFrom(
                          padding: const EdgeInsets.symmetric(vertical: 12),
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(8),
                          ),
                        ),
                      ),
                    ),
                    const SizedBox(width: 8),
                    Expanded(
                      child: OutlinedButton.icon(
                        onPressed:
                            _isLoading ? null : _selectScheduledOutputDir,
                        icon: const Icon(Icons.folder_open),
                        label: Text(
                          _scheduledOutputDirPath != null
                              ? '目标目录：${_scheduledOutputDirPath!.split('/').last}'
                              : '选择目标目录',
                          overflow: TextOverflow.ellipsis,
                        ),
                        style: OutlinedButton.styleFrom(
                          padding: const EdgeInsets.symmetric(vertical: 12),
                          shape: RoundedRectangleBorder(
                            borderRadius: BorderRadius.circular(8),
                          ),
                        ),
                      ),
                    ),
                  ],
                ),
                const SizedBox(height: 16),
                ElevatedButton.icon(
                  onPressed: _isLoading ? null : _executeBatchUnlock,
                  icon: const Icon(Icons.play_arrow_rounded),
                  label: Text(
                    '立即批量解锁'
                    '${_batchExecutionCount > 0 ? '（已执行 $_batchExecutionCount 次）' : ''}',
                  ),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: successGreen,
                    foregroundColor: Colors.white,
                    padding: const EdgeInsets.symmetric(vertical: 14),
                    shape: RoundedRectangleBorder(
                      borderRadius: BorderRadius.circular(10),
                    ),
                  ),
                ),
                const SizedBox(height: 16),
                const Divider(),
                const SizedBox(height: 8),
                Row(
                  children: [
                    const Icon(Icons.schedule, color: brandPink),
                    const SizedBox(width: 8),
                    Text(
                      '定时自动解锁',
                      style: theme.textTheme.titleMedium?.copyWith(
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                    const Spacer(),
                    Switch(
                      value: _scheduleEnabled,
                      onChanged:
                          _isLoading ? null : (v) => _handleScheduleToggle(v),
                      activeColor: brandPink,
                    ),
                  ],
                ),
                if (_scheduleEnabled) ...[
                  const SizedBox(height: 12),
                  // 时间选择按钮
                  OutlinedButton.icon(
                    onPressed: _isLoading ? null : _selectScheduledDateTime,
                    icon: const Icon(Icons.access_time),
                    label: Text(
                      _scheduledDateTime != null
                          ? '结束时间：${_scheduledDateTime!.year}-${_scheduledDateTime!.month.toString().padLeft(2, '0')}-${_scheduledDateTime!.day.toString().padLeft(2, '0')} ${_scheduledDateTime!.hour.toString().padLeft(2, '0')}:${_scheduledDateTime!.minute.toString().padLeft(2, '0')}'
                          : '选择结束时间',
                    ),
                    style: OutlinedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(
                          vertical: 12, horizontal: 16),
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(8),
                      ),
                      side: const BorderSide(color: brandPink),
                    ),
                  ),
                  const SizedBox(height: 12),
                  Container(
                    width: double.infinity,
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: brandPink.withOpacity(0.08),
                      borderRadius: BorderRadius.circular(8),
                    ),
                    child: Text(
                      _scheduledDateTime != null
                          ? '定时任务将持续运行至 ${_scheduledDateTime!.year}-${_scheduledDateTime!.month.toString().padLeft(2, '0')}-${_scheduledDateTime!.day.toString().padLeft(2, '0')} ${_scheduledDateTime!.hour.toString().padLeft(2, '0')}:${_scheduledDateTime!.minute.toString().padLeft(2, '0')}，期间每 $_checkIntervalMinutes 分钟检查一次新文件并自动解锁。'
                          : '请选择定时结束时间（不能选择过去的时间）',
                      style: theme.textTheme.bodySmall?.copyWith(
                        color: brandPink,
                      ),
                    ),
                  ),
                ] else ...[
                  // 定时任务未启用时，也显示时间选择按钮
                  const SizedBox(height: 12),
                  OutlinedButton.icon(
                    onPressed: _isLoading ? null : _selectScheduledDateTime,
                    icon: const Icon(Icons.access_time),
                    label: Text(
                      _scheduledDateTime != null
                          ? '已选择：${_scheduledDateTime!.year}-${_scheduledDateTime!.month.toString().padLeft(2, '0')}-${_scheduledDateTime!.day.toString().padLeft(2, '0')} ${_scheduledDateTime!.hour.toString().padLeft(2, '0')}:${_scheduledDateTime!.minute.toString().padLeft(2, '0')}'
                          : '选择结束时间',
                    ),
                    style: OutlinedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(
                          vertical: 12, horizontal: 16),
                      shape: RoundedRectangleBorder(
                        borderRadius: BorderRadius.circular(8),
                      ),
                    ),
                  ),
                  if (_scheduledDateTime != null) ...[
                    const SizedBox(height: 8),
                    Text(
                      '提示：开启定时开关后，将在指定时间执行批量解锁',
                      style: theme.textTheme.bodySmall?.copyWith(
                        color: Colors.grey[600],
                      ),
                    ),
                  ],
                ],
              ],
            ),
          ),
        ),
      ],
      const SizedBox(height: 32),

      // 加载指示器
      if (_isLoading && _progressMessage != null)
        Padding(
          padding: const EdgeInsets.only(top: 24),
          child: Column(
            children: [
              const CircularProgressIndicator(),
              const SizedBox(height: 16),
              Text(
                _progressMessage!,
                style: theme.textTheme.bodyMedium?.copyWith(
                  color: brandNavy,
                ),
                textAlign: TextAlign.center,
              ),
            ],
          ),
        ),
    ];
  }

  /// 构建匹配模式的内容
  List<Widget> _buildMatchModeContent(ThemeData theme) {
    // 自动折叠：如果输出目录已设置，默认折叠
    final shouldAutoCollapse =
        _outputDirExcel != null && _outputDirExcel!.isNotEmpty;

    return [
      // 两个拖拽区域 - 移到最上方
      Row(
        children: [
          // 草稿文件拖拽区域
          Expanded(
            child: _buildDraftFileDropZone(theme),
          ),
          const SizedBox(width: 16),
          // 目标文件拖拽区域
          Expanded(
            child: _buildTargetFileDropZone(theme),
          ),
        ],
      ),

      const SizedBox(height: 24),

      // 输出目录设置
      Card(
        elevation: 2,
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(12),
        ),
        child: InkWell(
          onTap: shouldAutoCollapse
              ? () {
                  setState(() {
                    _showOutputDirSettings = !_showOutputDirSettings;
                  });
                }
              : null,
          borderRadius: BorderRadius.circular(12),
          child: Padding(
            padding: const EdgeInsets.all(16),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    const Icon(Icons.folder_open, color: brandNavy),
                    const SizedBox(width: 8),
                    Expanded(
                      child: Text(
                        '输出目录',
                        style: theme.textTheme.titleMedium?.copyWith(
                          fontWeight: FontWeight.bold,
                        ),
                      ),
                    ),
                    if (shouldAutoCollapse)
                      Icon(
                        _showOutputDirSettings
                            ? Icons.keyboard_arrow_down
                            : Icons.keyboard_arrow_right,
                        color: brandNavy,
                      ),
                  ],
                ),
                if (!shouldAutoCollapse || _showOutputDirSettings) ...[
                  const SizedBox(height: 12),
                  OutlinedButton.icon(
                    onPressed: _isLoading ? null : _selectExcelOutputDir,
                    icon: const Icon(Icons.folder_open),
                    label: Text(
                      _outputDirExcel != null
                          ? _outputDirExcel!.split('/').last
                          : '选择输出目录',
                      overflow: TextOverflow.ellipsis,
                    ),
                    style: OutlinedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(vertical: 12),
                      minimumSize: const Size(double.infinity, 48),
                    ),
                  ),
                  if (_outputDirExcel == null)
                    Padding(
                      padding: const EdgeInsets.only(top: 8),
                      child: Text(
                        '⚠️ 请选择输出目录',
                        style: theme.textTheme.bodySmall?.copyWith(
                          color: Colors.orange[700],
                          fontSize: 11,
                        ),
                      ),
                    ),
                ] else ...[
                  const SizedBox(height: 8),
                  Text(
                    _outputDirExcel != null
                        ? _outputDirExcel!.split('/').last
                        : '未设置',
                    style: theme.textTheme.bodySmall?.copyWith(
                      color: Colors.grey[600],
                    ),
                  ),
                ],
              ],
            ),
          ),
        ),
      ),

      const SizedBox(height: 24),

      // 加载指示器
      if (_isLoading && _progressMessage != null)
        Padding(
          padding: const EdgeInsets.only(top: 24),
          child: Column(
            children: [
              const CircularProgressIndicator(),
              const SizedBox(height: 16),
              Text(
                _progressMessage!,
                style: theme.textTheme.bodyMedium?.copyWith(
                  color: brandNavy,
                ),
                textAlign: TextAlign.center,
              ),
            ],
          ),
        ),
    ];
  }

  /// 构建草稿文件拖拽区域
  Widget _buildDraftFileDropZone(ThemeData theme) {
    return DropTarget(
      onDragEntered: (_) => setState(() => _draftDragging = true),
      onDragExited: (_) => setState(() => _draftDragging = false),
      onDragDone: (details) async {
        setState(() => _draftDragging = false);
        if (details.files.isEmpty) return;
        final path = details.files.first.path;
        final file = File(path);
        if (await file.exists()) {
          final ext = path.toLowerCase().split('.').last;
          if (ext == 'xlsx') {
            setState(() {
              _draftFile = file;
              _statusMessage = '已选择草稿文件：${file.uri.pathSegments.last}';
            });
            // 检查输出目录，如果未设置则提示
            if (_outputDirExcel == null || _outputDirExcel!.isEmpty) {
              setState(() {
                _statusMessage = '请先选择输出目录';
              });
              return;
            }
            // 自动执行匹配（如果两个文件都已选择）
            _checkAndAutoMatch();
          } else {
            setState(() {
              _statusMessage = '草稿文件必须是 .xlsx 格式';
            });
          }
        }
      },
      child: Container(
        height: 320,
        decoration: BoxDecoration(
          color: _draftDragging
              ? successGreen.withOpacity(0.15)
              : (_draftFile != null
                  ? successGreen.withOpacity(0.1)
                  : brandCream.withOpacity(0.3)),
          borderRadius: BorderRadius.circular(16),
          border: Border.all(
            color: _draftDragging
                ? successGreen
                : (_draftFile != null
                    ? successGreen
                    : brandNavy.withOpacity(0.3)),
            width: _draftDragging ? 3 : 2,
          ),
        ),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(
              _draftFile != null ? Icons.check_circle : Icons.description,
              size: 64,
              color: _draftFile != null ? successGreen : brandNavy,
            ),
            const SizedBox(height: 16),
            Text(
              _draftFile != null ? '草稿文件已选择' : '拖拽草稿文件到这里',
              style: theme.textTheme.titleMedium?.copyWith(
                color: brandNavy,
                fontWeight: FontWeight.bold,
              ),
            ),
            if (_draftFile != null) ...[
              const SizedBox(height: 8),
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 16),
                child: Text(
                  _draftFile!.uri.pathSegments.last,
                  style: theme.textTheme.bodySmall?.copyWith(
                    color: Colors.grey[600],
                  ),
                  textAlign: TextAlign.center,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                ),
              ),
            ] else ...[
              const SizedBox(height: 8),
              Text(
                '支持 .xlsx 格式',
                style: theme.textTheme.bodySmall?.copyWith(
                  color: Colors.grey[600],
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }

  /// 构建目标文件拖拽区域
  Widget _buildTargetFileDropZone(ThemeData theme) {
    return DropTarget(
      onDragEntered: (_) => setState(() => _targetDragging = true),
      onDragExited: (_) => setState(() => _targetDragging = false),
      onDragDone: (details) async {
        setState(() => _targetDragging = false);
        if (details.files.isEmpty) return;
        final path = details.files.first.path;
        final file = File(path);
        if (await file.exists()) {
          final ext = path.toLowerCase().split('.').last;
          if (ext == 'xlsx') {
            setState(() {
              _targetFile = file;
              _statusMessage = '已选择目标文件：${file.uri.pathSegments.last}';
            });
            // 检查输出目录，如果未设置则提示
            if (_outputDirExcel == null || _outputDirExcel!.isEmpty) {
              setState(() {
                _statusMessage = '请先选择输出目录';
              });
              return;
            }
            // 自动执行匹配（如果两个文件都已选择）
            _checkAndAutoMatch();
          } else {
            setState(() {
              _statusMessage = '目标文件必须是 .xlsx 格式';
            });
          }
        }
      },
      child: Container(
        height: 320,
        decoration: BoxDecoration(
          color: _targetDragging
              ? successGreen.withOpacity(0.15)
              : (_targetFile != null
                  ? successGreen.withOpacity(0.1)
                  : brandCream.withOpacity(0.3)),
          borderRadius: BorderRadius.circular(16),
          border: Border.all(
            color: _targetDragging
                ? successGreen
                : (_targetFile != null
                    ? successGreen
                    : brandNavy.withOpacity(0.3)),
            width: _targetDragging ? 3 : 2,
          ),
        ),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(
              _targetFile != null ? Icons.check_circle : Icons.folder,
              size: 64,
              color: _targetFile != null ? successGreen : brandNavy,
            ),
            const SizedBox(height: 16),
            Text(
              _targetFile != null ? '目标文件已选择' : '拖拽目标文件到这里',
              style: theme.textTheme.titleMedium?.copyWith(
                color: brandNavy,
                fontWeight: FontWeight.bold,
              ),
            ),
            if (_targetFile != null) ...[
              const SizedBox(height: 8),
              Padding(
                padding: const EdgeInsets.symmetric(horizontal: 16),
                child: Text(
                  _targetFile!.uri.pathSegments.last,
                  style: theme.textTheme.bodySmall?.copyWith(
                    color: Colors.grey[600],
                  ),
                  textAlign: TextAlign.center,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                ),
              ),
            ] else ...[
              const SizedBox(height: 8),
              Text(
                '支持 .xlsx 格式',
                style: theme.textTheme.bodySmall?.copyWith(
                  color: Colors.grey[600],
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }

  /// 检查并自动执行匹配
  Future<void> _checkAndAutoMatch() async {
    // 如果正在匹配，不重复执行
    if (_isAutoMatching || _isLoading) return;

    // 检查必要条件
    if (_draftFile == null || _targetFile == null) return;
    if (_outputDirExcel == null || _outputDirExcel!.isEmpty) {
      setState(() {
        _statusMessage = '请先选择输出目录';
      });
      return;
    }

    // 自动执行匹配，使用 try-finally 确保状态重置
    _isAutoMatching = true;
    try {
      await _executeMatch();
    } finally {
      _isAutoMatching = false;
    }
  }

  /// 选择 Excel 输出目录
  Future<void> _selectExcelOutputDir() async {
    final directory = await getDirectoryPath();
    if (directory != null && directory.isNotEmpty) {
      setState(() {
        _outputDirExcel = directory;
      });
      await _persistBatchPreferences();
      // 如果文件已选择，自动执行匹配
      _checkAndAutoMatch();
    }
  }

  /// 清理匹配状态（清空文件、重置加载状态）
  void _clearMatchState({String? keepStatusMessage}) {
    setState(() {
      _draftFile = null;
      _targetFile = null;
      _isLoading = false;
      _isAutoMatching = false;
      _progressMessage = null;
      // 保留状态消息（如果有），否则清空
      if (keepStatusMessage != null) {
        _statusMessage = keepStatusMessage;
      } else if (_statusMessage != null && _statusMessage!.contains('匹配')) {
        // 如果是匹配相关的消息，保留
      } else {
        _statusMessage = null;
      }
    });
  }

  /// 执行匹配
  Future<void> _executeMatch() async {
    if (_draftFile == null || _targetFile == null) {
      await _showInfoDialog('缺少文件', '请先选择草稿文件和目标文件。');
      return;
    }

    // 检查输出目录（提前检查）
    if (_outputDirExcel == null || _outputDirExcel!.isEmpty) {
      _updateLoadingState(
        isLoading: false,
        statusMessage: '请先选择输出目录',
        progressMessage: null,
      );
      return;
    }

    _updateLoadingState(
      isLoading: true,
      statusMessage: null,
      progressMessage: '正在匹配文件...',
    );

    String? finalStatusMessage;
    try {
      // 确定输出文件路径（必须使用已设置的输出目录）
      final outputDir = _outputDirExcel!;
      final targetFileName = _targetFile!.uri.pathSegments.last;
      final baseName = targetFileName.replaceAll('.xlsx', '');
      final timestamp = DateTime.now().millisecondsSinceEpoch;
      final outputPath = '$outputDir/${baseName}_matched_$timestamp.xlsx';

      // 执行匹配（自动识别列结构，不再需要 targetColumns 参数）
      final result = await _matchService.matchExcelFiles(
        draftFile: _draftFile!,
        targetFile: _targetFile!,
        outputPath: outputPath,
      );

      if (result.success) {
        finalStatusMessage =
            '${result.message}\n输出文件：${outputPath.split('/').last}';
      } else {
        finalStatusMessage = result.message;
      }
    } catch (e, st) {
      debugPrint('[ExcelEditorPage] 匹配失败: $e');
      debugPrint('[ExcelEditorPage] 堆栈: $st');
      finalStatusMessage = '匹配失败：$e';
    } finally {
      // 无论成功还是失败，都清理匹配状态
      _clearMatchState(keepStatusMessage: finalStatusMessage);
    }
  }
}
