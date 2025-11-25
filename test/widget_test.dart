import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:myexcle/main.dart';

void main() {
  testWidgets('App shell renders main screen', (WidgetTester tester) async {
    await tester.pumpWidget(const MyApp());

    expect(find.text('文件解锁工具'), findsOneWidget);
    expect(find.byType(AppBar), findsOneWidget);
  });
}
