# EPF Unlocker

[![CI](https://github.com/lnuxe/EPF-Unlocker/actions/workflows/ci.yml/badge.svg)](https://github.com/lnuxe/EPF-Unlocker/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)

> **项目地址**：[https://github.com/lnuxe/EPF-Unlocker](https://github.com/lnuxe/EPF-Unlocker)

一个用于解锁 Excel (.xlsx) 和 PDF 文件保护的工具，支持单文件解锁、批量解锁和定时自动解锁功能。

## 快速开始

```bash
# 克隆项目
git clone https://github.com/lnuxe/EPF-Unlocker.git
cd EPF-Unlocker

# 安装依赖
flutter pub get

# 运行应用
flutter run
```

## 功能特性

### 核心功能
- ✅ **单文件解锁**：支持拖拽或选择文件进行解锁
- ✅ **批量解锁**：支持递归扫描目录下所有文件并批量解锁
- ✅ **定时自动解锁**：设置定时任务，自动监控目录并解锁新文件
- ✅ **自动保存**：解锁后自动保存到指定目录（如果已设置批量解锁目标目录）

### 支持的文件格式
- **Excel 文件**：`.xlsx` 格式（移除工作表保护）
- **PDF 文件**：`.pdf` 格式（移除密码保护）
- ⚠️ **不支持**：`.xls` 格式（旧版 Excel 格式）

> **注意**：如果遇到 `.xls` 文件，程序会提示使用 Microsoft Excel 将其转换为 `.xlsx` 格式后再解锁。

## 使用方法

### 1. 单文件解锁

#### 方法一：拖拽文件
- 直接将文件拖拽到应用窗口的拖拽区域
- 文件会自动解锁并保存（如果已设置批量解锁目标目录）
- 或者会弹出保存对话框让您选择保存位置

#### 方法二：选择文件
- 点击"选择文件解锁"按钮
- 在弹出的文件选择对话框中选择要解锁的文件
- 选择保存位置

### 2. 批量解锁

1. **点击右上角设置按钮**（齿轮图标）展开批量解锁设置
2. **选择源目录**：点击"选择源目录"按钮，选择包含待解锁文件的目录
3. **选择目标目录**：点击"选择目标目录"按钮，选择解锁后文件的保存位置
4. **执行批量解锁**：
   - 点击"立即批量解锁"按钮立即执行一次
   - 或开启"定时自动解锁"开关，设置间隔时间后自动执行

### 3. 定时自动解锁

1. 先设置源目录和目标目录
2. 开启"定时自动解锁"开关
3. 调整"间隔时间"滑块，设置自动解锁的间隔（1-120 分钟）
4. 定时任务启动后会：
   - **立即执行一次**解锁检查（检查是否有新的待解锁文件）
   - 然后按设置的间隔时间定期执行

> **注意**：拖拽定时时间滑块时不会触发解锁，只有释放滑块时才会更新定时器。

## 平台配置

### Windows 平台配置

为了打包为 `.exe` 文件并正常运行，Windows 平台需要以下配置：

#### 1. 文件访问权限配置

文件：`windows/runner/runner.exe.manifest`

当前配置已包含基本权限：
- ✅ **执行级别**：`asInvoker`（以当前用户权限运行，不需要管理员权限）
- ✅ **兼容性**：支持 Windows 10 和 Windows 11
- ✅ **DPI 感知**：`PerMonitorV2`（支持多显示器高DPI）
- ✅ **文件系统访问**：已启用 `broadFileSystemAccess`（允许访问用户选择的文件和目录）

**重要说明 - UUID 的作用**：

manifest 文件中的 UUID `{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}` 是 **Windows 10/11 的系统标识符**，不是应用程序特定的 UUID。

- ✅ **这个 UUID 是固定的**：它是 Windows 10/11 的操作系统标识符，所有 Windows 应用都使用相同的值
- ✅ **不需要修改**：即使在其他电脑上重新打包，也不需要更改这个 UUID
- ✅ **不会影响权限**：文件权限由 `requestedExecutionLevel` 和 `broadFileSystemAccess` 决定，而不是 UUID
- ✅ **可以安全打包**：在任何电脑上打包为 .exe 文件，权限配置都会正常工作

**权限配置说明**：

当前已启用 `broadFileSystemAccess` 权限：
```xml
<capabilities>
  <rescap:Capability Name="broadFileSystemAccess" 
    xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities" />
</capabilities>
```

这个权限允许应用访问用户通过文件选择器选择的文件和目录，满足批量解锁功能的需求。

#### 2. 应用图标配置

Windows 应用图标文件位置：`windows/runner/resources/app_icon.ico`

**如何设置自定义图标**：

1. **准备图标文件**：
   - 将图片转换为 `.ico` 格式（可以使用在线工具或 ImageMagick）
   - 建议尺寸：256x256 像素或更大（支持多尺寸）
   - 文件路径：`windows/runner/resources/app_icon.ico`

2. **替换图标文件**：
   - 直接替换 `windows/runner/resources/app_icon.ico` 文件
   - 确保文件名完全一致（`app_icon.ico`）

3. **重新构建**：
   ```bash
   flutter clean
   flutter build windows --release
   ```

**图标文件引用**：

图标在 `windows/runner/Runner.rc` 文件中引用：
```rc
IDI_APP_ICON            ICON                    "resources\\app_icon.ico"
```

只需替换 `app_icon.ico` 文件，重新构建后图标就会更新。

#### 3. 打包为 EXE 文件

使用以下命令打包 Windows 应用：

```bash
# 开发版构建
flutter build windows

# 发布版构建（优化）
flutter build windows --release
```

打包后的可执行文件位于：`build/windows/runner/Release/myexcle.exe`

#### 4. 分发应用

**方法一：直接分发文件夹**
- 将 `build/windows/runner/Release/` 文件夹中的所有文件一起分发
- 包括：`myexcle.exe` 和相关的 DLL 文件

**方法二：创建安装程序**
- 使用 Inno Setup、NSIS 等工具创建安装程序
- 确保包含所有依赖项

### macOS 平台配置

#### 权限配置

文件：`macos/Runner/DebugProfile.entitlements` 和 `macos/Runner/Release.entitlements`

已配置权限：
- ✅ **文件访问**：`com.apple.security.files.user-selected.read-write`（允许访问用户选择的文件）
- ✅ **沙箱**：`com.apple.security.app-sandbox`（应用沙箱）

#### 打包为 APP 文件

```bash
# 开发版构建
flutter build macos

# 发布版构建
flutter build macos --release
```

打包后的应用位于：`build/macos/Build/Products/Release/myexcle.app`

## 依赖项说明

### 核心依赖
- `archive: ^3.6.1` - ZIP/归档文件处理（用于 .xlsx 文件）
- `xml: ^6.5.0` - XML 解析和修改（用于移除 Excel 保护）
- `excel: ^4.0.6` - Excel 文件读写（仅支持 .xlsx）
- `syncfusion_flutter_pdf: ^27.2.5` - PDF 文件处理和解锁
- `file_selector: ^1.0.3` - 桌面文件选择对话框
- `file_picker: ^8.0.0` - 跨平台文件选择
- `desktop_drop: ^0.4.0` - 桌面文件拖拽支持
- `shared_preferences: ^2.2.2` - 本地数据存储（保存批量解锁设置）

### PDF 解锁外部依赖

PDF 解锁功能需要以下工具之一（macOS/Windows）：

**方法一：qpdf**
```bash
# macOS
brew install qpdf

# Windows
# 从 https://qpdf.sourceforge.io/ 下载安装
```

**方法二：Python + pypdf**
```bash
pip install pypdf
```

> **注意**：如果两者都不可用，PDF 解锁功能将无法使用，会显示相应提示。

## 技术实现

### Excel 解锁原理
1. 将 `.xlsx` 文件作为 ZIP 归档解压
2. 查找所有工作表的 XML 文件（`xl/worksheets/*.xml`）
3. 从 XML 中移除 `<sheetProtection>` 标签
4. 重新打包为新的 `.xlsx` 文件

### PDF 解锁原理
- 使用 `syncfusion_flutter_pdf` 库读取 PDF
- 移除安全限制和密码保护
- 生成新的无保护 PDF 文件

### 批量解锁流程
1. 递归扫描源目录下的所有 `.xlsx` 和 `.pdf` 文件
2. 逐个处理每个文件
3. 保存到目标目录，如果文件已存在则添加时间戳

### 定时任务实现
- 使用 `Timer` 实现在指定时间执行一次
- 可以选择今天或未来10天内的任何时间
- 系统会自动验证时间不能是过去的时间
- 执行完成后自动停止，如需重复执行需要重新设置

## 常见问题

### Q: 为什么 `.xls` 文件不支持？
A: `.xls` 是旧版 Excel 格式（二进制格式），需要专门的库解析。当前项目已移除 `.xls` 支持，建议使用 Microsoft Excel 将 `.xls` 文件另存为 `.xlsx` 格式。

### Q: 文件已损坏错误怎么办？
A: 如果文件多次被处理或已损坏，建议：
1. 使用原始未解锁的文件
2. 使用 Microsoft Excel 打开并另存为新的 `.xlsx` 文件
3. 然后再尝试解锁

### Q: Windows 上无法保存文件？
A: 检查以下内容：
1. 确保目标目录存在
2. 确保应用有权限访问该目录
3. 如果使用批量解锁目标目录，确保目录路径正确
4. 可以尝试重新选择保存位置

### Q: 定时任务不执行？
A: 检查：
1. 是否已设置源目录和目标目录
2. 是否已选择定时执行时间（不能是过去的时间）
3. 定时开关是否已开启
4. 选择的执行时间是否在未来（系统会自动验证）

### Q: Windows manifest 中的 UUID 是否唯一？在其他电脑打包会不会影响权限？
A: **UUID `{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}` 不是应用特定的**：
- ✅ 这是 Windows 10/11 的系统标识符，所有应用都使用相同的值
- ✅ 不需要修改，在任何电脑上打包都使用相同的 UUID
- ✅ 不会影响文件权限，权限由 `requestedExecutionLevel` 和 `broadFileSystemAccess` 决定
- ✅ 可以安全地在其他电脑上重新打包，权限配置会正常工作

## 开发说明

### 项目结构
```
lib/
├── main.dart                 # 应用入口
├── models/                   # 数据模型
│   └── excel_models.dart
├── services/                 # 业务逻辑服务
│   ├── excel_file_service.dart      # 文件选择和解锁
│   ├── excel_parse_service.dart     # Excel 解析
│   ├── excel_export_service.dart    # Excel 导出
│   ├── pdf_unlock_service.dart      # PDF 解锁
│   └── xml_preserve_export_service.dart  # XML 保持格式导出
└── ui/                       # 用户界面
    ├── app_theme.dart        # 主题配置
    └── excel_editor_page.dart # 主界面
```

### 代码规范
- 遵循 Flutter/Dart 代码规范
- 使用 `debugPrint` 进行调试日志输出
- 完善的错误处理和用户提示
- 提交前运行 `flutter format .` 和 `flutter analyze`

### 本地开发

```bash
# 克隆项目
git clone https://github.com/lnuxe/EPF-Unlocker.git
cd EPF-Unlocker

# 安装依赖
flutter pub get

# 运行代码分析
flutter analyze

# 运行测试
flutter test

# 运行应用（需要先选择平台）
flutter run -d windows  # Windows
flutter run -d macos    # macOS
flutter run -d chrome   # Web
```

## 开源计划

### 项目目标

EPF Unlocker 是一个开源项目，旨在为社区提供一个简单、高效的文件解锁工具。我们致力于：

- 🎯 **持续改进**：根据用户反馈不断优化功能和性能
- 🔒 **安全可靠**：确保文件处理过程安全，不泄露用户数据
- 🌍 **跨平台支持**：支持 Windows、macOS 和 Web 平台
- 📚 **文档完善**：提供详细的使用文档和开发指南

### 版本规划

- **v1.0.x**：基础功能稳定版本
  - ✅ Excel (.xlsx) 文件解锁
  - ✅ PDF 文件解锁
  - ✅ 批量解锁功能
  - ✅ 定时自动解锁

- **v1.1.x**（计划中）：
  - 🔄 性能优化（并发处理）
  - 🔄 UI/UX 改进
  - 🔄 错误处理增强

- **v1.2.x**（未来）：
  - 📋 支持更多文件格式
  - 📋 命令行工具支持
  - 📋 插件系统

### 贡献指南

我们欢迎所有形式的贡献！无论是代码、文档、测试还是问题反馈，都是对项目的宝贵支持。

#### 如何贡献

1. **Fork 项目**
   ```bash
   # Fork 仓库后，克隆到本地
   git clone https://github.com/你的用户名/EPF-Unlocker.git
   cd EPF-Unlocker
   ```

2. **创建功能分支**
   ```bash
   git checkout -b feature/你的功能名称
   # 或
   git checkout -b fix/修复的问题描述
   ```

3. **开发与测试**
   ```bash
   # 安装依赖
   flutter pub get
   
   # 运行代码分析
   flutter analyze
   
   # 运行测试
   flutter test
   
   # 确保代码通过 CI 检查
   ```

4. **提交代码**
   ```bash
   # 提交前确保代码格式化
   flutter format .
   
   # 提交更改
   git add .
   git commit -m "feat: 添加新功能描述"
   # 或
   git commit -m "fix: 修复问题描述"
   ```

5. **推送并创建 Pull Request**
   ```bash
   git push origin feature/你的功能名称
   ```
   然后在 GitHub 上创建 Pull Request，详细描述你的更改。

#### 代码规范

- 遵循 [Flutter/Dart 代码规范](https://dart.dev/guides/language/effective-dart/style)
- 使用有意义的变量和函数名
- 添加必要的注释，特别是复杂逻辑
- 确保所有测试通过
- 代码提交前运行 `flutter format .` 和 `flutter analyze`

#### 提交信息规范

我们使用 [Conventional Commits](https://www.conventionalcommits.org/) 规范：

- `feat:` 新功能
- `fix:` 修复 bug
- `docs:` 文档更新
- `style:` 代码格式调整（不影响功能）
- `refactor:` 代码重构
- `test:` 测试相关
- `chore:` 构建过程或辅助工具的变动

示例：
```
feat: 添加批量解锁并发处理功能
fix: 修复 Windows 平台文件保存权限问题
docs: 更新 README 中的使用说明
```

#### Issue 报告

如果发现 bug 或有功能建议，请创建 Issue：

- **Bug 报告**：请提供详细的复现步骤、预期行为和实际行为
- **功能建议**：请描述功能的使用场景和预期效果
- **问题讨论**：欢迎在 Issue 中讨论技术问题和实现方案

#### 行为准则

- 尊重所有贡献者
- 接受建设性的批评
- 专注于对项目最有利的事情
- 对其他社区成员表示同理心

### CI/CD 流程

项目使用 GitHub Actions 进行持续集成，每次推送代码或创建 Pull Request 时都会自动运行：

1. **代码分析** (`flutter analyze`)：检查代码质量和规范
2. **单元测试** (`flutter test`)：运行所有测试用例
3. **构建验证**：验证 Web、Windows、macOS 平台的构建

CI 状态徽章显示在 README 顶部，绿色表示所有检查通过。

### 维护者

当前维护者：[@lnuxe](https://github.com/lnuxe)

### 许可证

本项目采用 [Apache-2.0](LICENSE) 许可证开源。

## 更新日志

### v1.0.0+1
- ✅ 实现 Excel (.xlsx) 文件解锁功能
- ✅ 实现 PDF 文件解锁功能
- ✅ 实现批量解锁功能
- ✅ 实现定时自动解锁功能
- ✅ 实现递归目录扫描
- ✅ 实现拖拽文件解锁
- ✅ 优化 UI，默认简洁视图
- ✅ 移除 .xls 格式支持
- ✅ 修复文件保存权限问题
- ✅ 优化定时任务逻辑

