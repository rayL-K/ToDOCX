# ToDOCX - 智能排版工具

琐碎排版，一键告别。

一个基于 PyQt 的桌面工具，将 DOCX / LaTeX(.tex) / Markdown(.md) 统一转换为排版良好的 DOCX 文档，并支持自定义样式模板。

## 功能

1. **DOCX 智能再排版**  
   加载现有 Word 文档，对段落类型进行识别与调整（标题、正文、代码、表格、公式等），一键应用左侧样式设置重新生成 DOCX。

2. **LaTeX(.tex) → DOCX**  
   - 识别章节标题并自动编号（如「一、」「1.」「1.1」等）  
   - 将表格环境解析为真实的 Word 表格，并在表格下方生成「表1  标题」样式的图表标题  
   - 将 `lstlisting` 等代码环境保留为等宽字体代码块，并在下方生成「代码1  标题」说明  
   - 将块级公式转换为 Word 公式对象（OMML），行内 `$...$` 渲染为 Cambria Math 斜体  
   - 正确处理 LaTeX 转义（如 `\_`, `\%`, `\&`）和 URL（`\url{}` / `\href{}`）。

3. **Markdown(.md) → DOCX**  
   通过智能排版页面选择 `.md` 文件，按标题层级、正文、列表等结构生成 DOCX，并可复用同一套样式设置和模板。

4. **样式模板管理**  
   - 内置「默认样式」模板  
   - 支持保存当前样式为**自定义模板**  
   - 支持对用户模板进行**重命名**和删除

> 说明：原有 PDF 转 Word 和 AI 排版功能已移除，当前版本专注于 DOCX/LaTeX/Markdown → DOCX 的智能排版。

## 安装

项目使用 `uv` 管理依赖，建议步骤如下：

```bash
uv sync
uv pip install PyQt5==5.15.9
```

确保在 Windows 上使用本项目自带的虚拟环境：`.venv`。

## 运行（开发环境）

在项目根目录执行：

```bash
.\.venv\Scripts\python.exe main.py
```

## 使用说明

1. 启动程序后，顶部显示「ToDOCX」与副标题「琐碎排版，一键告别」。
2. 在中间区域的「文件拖放区域」中，拖入或选择：
   - `.docx`：对现有 Word 进行智能再排版；
   - `.tex`：将 LaTeX 文档转换为排版良好的 DOCX；
   - `.md` / `.markdown`：将 Markdown 文档转换为 DOCX。
3. 左侧「样式」标签中调整：标题、正文、图表标题、代码等样式（字体、中英文字体、字号、行距、缩进）。
4. 「模板」标签中可以：
   - 加载默认或自定义模板；
   - 保存当前样式为新模板；
   - 对自定义模板进行重命名或删除。
5. 右侧预览区中，可查看段落类型识别结果，右键修改段落类型。
6. 点击底部的「开始转换」按钮，生成新的 DOCX 文件。

## 打包为 Windows 可执行文件（EXE）

推荐使用 PyInstaller 进行打包（在虚拟环境中执行）：

1. 安装 PyInstaller：

```bash
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

2. 在项目根目录打包：

```bash
.\.venv\Scripts\pyinstaller --name ToDOCX \
    --windowed \
    --onefile \
    main.py
```

执行成功后，在 `dist/` 目录下会生成一个 `ToDOCX.exe`，可直接分发给其他 Windows 用户使用（前提是系统已安装相应的字体）。
