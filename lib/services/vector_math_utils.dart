import 'dart:math' as math;
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
  /// 支持两层（如 "1.14"）和三层（如 "3.6.1"）格式
  /// 对于三层格式，l2 会包含第二层和第三层的信息：l2 * 100 + l3
  /// 例如 "3.6.1" -> l1=3, l2=6*100+1=601
  /// 这里为了简单，假定 level1 ∈ [0, 1000]，level2 ∈ [0, 100000]（支持三层时最大为 1000*100）
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
      final secondLevel = int.tryParse(parts[1]) ?? 0;
      if (parts.length > 2) {
        // 三层格式：如 "3.6.1" -> l1=3, l2=6*100+1=601
        final thirdLevel = int.tryParse(parts[2]) ?? 0;
        l2 = secondLevel * 100 + thirdLevel;
      } else {
        // 两层格式：如 "1.14" -> l1=1, l2=14
        l2 = secondLevel;
      }
    }

    // 简单归一化
    final l1Norm = l1 / 1000.0; // 假设最多 1000 大项
    final l2Norm = l2 / 100000.0; // 支持三层时，最大为 1000*100=100000

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
  /// w4: Qty 不匹配惩罚权重（提高权重以确保Qty匹配）
  static double computeMatchScore({
    required RowFeatureVector targetVec,
    required RowFeatureVector draftVec,
    required String targetDesc,
    required String draftDesc,
    required String targetUnit,
    required String? draftUnit,
    required double targetQty,
    required double? draftQty,
    double w1 = 0.40, // 降低Item权重
    double w2 = 0.35, // 降低描述权重
    double w3 = 0.10, // Unit权重保持不变
    double w4 = 0.15, // 提高Qty权重（从0.05提高到0.15）
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

    // 4. Qty 惩罚（比例差，如果Qty都有值且差异较大，给予更大惩罚）
    double qtyPenalty = 0.0;
    if (targetQty > 0 && (draftQty ?? 0) > 0) {
      final diff = (targetQty - (draftQty ?? 0)).abs();
      final maxQty = targetQty > (draftQty ?? 0) ? targetQty : (draftQty ?? 0);
      final relativeDiff = diff / maxQty;
      // 如果Qty差异超过20%，给予更大的惩罚
      if (relativeDiff > 0.2) {
        qtyPenalty = (relativeDiff * 2.0).clamp(0.0, 2.0); // 最大惩罚为2.0
      } else {
        qtyPenalty = relativeDiff.clamp(0.0, 1.0);
      }
    } else if (targetQty > 0 && (draftQty ?? 0) == 0) {
      // 目标有Qty但草稿没有，给予中等惩罚
      qtyPenalty = 0.3;
    } else if (targetQty == 0 && (draftQty ?? 0) > 0) {
      // 目标没有Qty但草稿有，给予较小惩罚（因为可以填充）
      qtyPenalty = 0.1;
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
    return sum == 0 ? 0.0 : math.sqrt(sum);
  }
}

/// 文本归一化
String _normalizeText(String text) {
  return text
      .toLowerCase()
      .replaceAll(RegExp(r'\s+'), '')
      .replaceAll(RegExp(r'[()（）]'), '');
}
