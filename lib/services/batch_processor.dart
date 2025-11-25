import 'dart:async';

/// 批量处理器：控制并发数量，管理任务队列
///
/// 用于批量解锁场景，避免同时处理过多文件导致系统资源耗尽和 UI 卡顿
/// 采用并发控制机制
class BatchProcessor {
  /// 默认最大并发数：同时处理 3 个文件
  /// 这个数量在保证效率的同时，避免资源过度消耗
  static const int defaultMaxConcurrency = 3;

  final int maxConcurrency;
  final _semaphore = StreamController<void>();
  bool _disposed = false;

  BatchProcessor({int? maxConcurrency})
      : maxConcurrency = maxConcurrency ?? defaultMaxConcurrency {
    // 初始化信号量，创建初始可用槽位
    for (var i = 0; i < this.maxConcurrency; i++) {
      if (!_disposed) {
        _semaphore.add(null);
      }
    }
  }

  /// 批量处理任务列表
  ///
  /// [items] 待处理的项目列表
  /// [processor] 处理单个项目的函数，返回处理结果
  /// [onProgress] 进度回调函数，每次处理完成后调用 (已完成数量, 总数量, 当前项)
  ///
  /// 返回处理结果列表，顺序与输入列表一致，但会过滤掉处理失败的项目
  Future<List<T>> processBatch<T, E>({
    required List<E> items,
    required Future<T> Function(E item, int index) processor,
    void Function(int completed, int total, E? current)? onProgress,
  }) async {
    if (items.isEmpty) return [];

    final results = <T?>[];
    int completedCount = 0;
    final lock = Lock();

    // 初始化结果列表
    for (var i = 0; i < items.length; i++) {
      results.add(null);
    }

    // 并发处理所有任务，但通过信号量控制并发数量
    final futures = <Future<void>>[];
    for (var i = 0; i < items.length; i++) {
      final index = i;
      final item = items[index];

      futures.add(_processItem(
        index: index,
        item: item,
        processor: processor,
        onResult: (result) async {
          await lock.synchronized(() async {
            results[index] = result;
            completedCount++;
          });
          if (onProgress != null && !_disposed) {
            onProgress(completedCount, items.length, item);
          }
        },
        onError: () async {
          // 错误处理：记录错误但不中断整个批次
          await lock.synchronized(() async {
            completedCount++;
          });
          if (onProgress != null && !_disposed) {
            onProgress(completedCount, items.length, item);
          }
        },
      ));
    }

    // 等待所有任务完成
    await Future.wait(futures);

    // 转换结果，过滤掉 null 值（由于错误导致的）
    return results.where((r) => r != null).map((r) => r as T).toList();
  }

  /// 处理单个任务项
  ///
  /// 内部使用信号量控制并发数量，确保同时运行的任务不超过 maxConcurrency
  Future<void> _processItem<T, E>({
    required int index,
    required E item,
    required Future<T> Function(E item, int index) processor,
    required Future<void> Function(T result) onResult,
    required Future<void> Function() onError,
  }) async {
    if (_disposed) return;

    // 等待可用槽位
    await _semaphore.stream.first;

    if (_disposed) return;

    try {
      final result = await processor(item, index);
      await onResult(result);
    } catch (e) {
      // 忽略错误，继续处理下一个
      await onError();
    } finally {
      // 释放槽位，允许下一个任务开始
      if (!_disposed && !_semaphore.isClosed) {
        _semaphore.add(null);
      }
    }
  }

  /// 释放资源
  void dispose() {
    _disposed = true;
    _semaphore.close();
  }
}

/// 简单的互斥锁，用于保护共享资源访问
class Lock {
  Future<void>? _waiting;
  bool _locked = false;

  /// 在锁保护下执行异步操作
  Future<T> synchronized<T>(Future<T> Function() action) async {
    // 等待之前的操作完成
    while (_locked) {
      final completer = Completer<void>();
      final previous = _waiting;
      _waiting = completer.future;
      if (previous != null) {
        await previous;
      }
      await completer.future;
    }

    _locked = true;
    try {
      return await action();
    } finally {
      _locked = false;
      final next = _waiting;
      _waiting = null;
      if (next != null) {
        // 通知等待的协程可以继续
      }
    }
  }
}
