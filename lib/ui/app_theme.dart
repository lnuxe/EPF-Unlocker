import 'package:flutter/material.dart';

// 主题色
const Color brandNavy = Color(0xFF0A1F44); // 深蓝主色（AppBar/深背景/主按钮）
const Color brandDarkBlue = Color(0xFF24479F); // 蓝色-按钮/强调色
const Color brandSteel = Color(0xFF5B84C0); // 浅蓝灰，次要卡片/辅助色
const Color brandCream = Color(0xFFE1E7F4); // 纸白/浅灰（背景/卡片/高亮）
const Color brandPink = Color(0xFF4361E4); // 浅蓝/标签/高亮
const Color brandRed = Color(0xFFE63946); // 红色（错误/警告/重要强调，保留）
const Color successGreen = Color(0xFF2E7D32); // 成功提示绿色，保留
const Color infoBlue = brandDarkBlue;

/// 统一背景色，较浅的灰色
const Color backgroundWhite = brandCream;

ThemeData buildAppTheme() {
  final base = ThemeData.light();
  return base.copyWith(
    scaffoldBackgroundColor: backgroundWhite,
    colorScheme: const ColorScheme(
      brightness: Brightness.light,
      primary: brandNavy,
      onPrimary: Colors.white,
      secondary: brandDarkBlue,
      onSecondary: Colors.white,
      surface: brandCream,
      onSurface: brandNavy,
      error: brandRed,
      onError: Colors.white,
    ),
    appBarTheme: const AppBarTheme(
      backgroundColor: brandNavy,
      foregroundColor: Colors.white,
      elevation: 2,
    ),
    textButtonTheme: TextButtonThemeData(
      style: TextButton.styleFrom(
        foregroundColor: Colors.white,
        backgroundColor: brandDarkBlue,
        padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(8),
        ),
      ),
    ),
    switchTheme: SwitchThemeData(
      thumbColor: WidgetStateProperty.resolveWith(
        (states) => states.contains(WidgetState.selected)
            ? brandSteel
            : Colors.grey.shade400,
      ),
      trackColor: WidgetStateProperty.resolveWith(
        (states) => states.contains(WidgetState.selected)
            ? brandNavy.withOpacity(0.5)
            : Colors.grey.shade300,
      ),
    ),
    cardColor: brandCream,
    dialogTheme: const DialogTheme(backgroundColor: brandCream),
  );
}
