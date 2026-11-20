# 文件转换助手

一个纯本地运行的桌面文件转换工具，支持 Word / PDF / 图片之间的互转，以及常用的 PDF 页面处理。界面基于 PySide6，转换全部在本机完成，**不联网、不上传文件**。

## 功能

### 格式转换

| 输入 | 输出 | 说明 |
| --- | --- | --- |
| Word（.docx/.doc/.rtf/.odt/.txt） | PDF | 调用本机 Microsoft Word 导出，排版最保真 |
| PDF | Word（.docx） | 基于 pdf2docx 按版式重建 |
| PDF | 图片（PNG/JPG/WebP/BMP） | 渲染 DPI 可调，支持指定页码 |
| 图片（JPG/PNG/BMP/GIF/TIFF/WebP） | PDF | 单张转换，或多张合并为一个 PDF |
| 图片 | 图片 | JPG/PNG/WebP/BMP 互转，可调质量与最长边 |

- 拖拽添加文件，支持批量队列、逐条状态与输出位置展示
- 页码支持 `1-3,5,8-` 写法（中英文逗号/分号、`至`/`到` 均可）
- 输出同名文件自动追加 `(1)`、`(2)`，不会覆盖已有文件
- 可选择「输出到原目录」或自定义输出目录

### PDF 工具箱

- **合并**：多个 PDF 合并为一个
- **拆分**：逐页拆分 / 每 N 页一个文件 / 自定义页码范围
- **旋转**：按页或全部页面旋转 90° / 180° / 270°
- **加密**：AES-256 加密，可设置打开密码与权限
- **解密**：输入密码去除加密
- **提取文字**：导出为 txt
- **提取图片**：导出 PDF 内嵌图片
- **压缩**：轻 / 标准 / 强三档

## 环境要求

- Windows 10 / 11
- Python 3.10+（开发使用 3.11）
- Word → PDF 功能需要本机安装 Microsoft Word；未安装时其余功能不受影响

## 从源码运行

```powershell
python -m pip install -r requirements.txt
python main.py
```

## 打包为 exe

```powershell
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean 文件转换助手.spec
```

产物位于 `dist\文件转换助手\`，整个文件夹拷贝即可分发，双击 `文件转换助手.exe` 运行。打包体积约 320MB（内置 Python 运行时与 Qt、PyMuPDF 等依赖）。

## 目录结构

```
FileConverter/
├── main.py                  # PySide6 界面、工作线程、程序入口
├── converters.py            # 转换核心：格式互转 + PDF 工具箱
├── theme.py                 # 全局 QSS 样式
├── requirements.txt         # 依赖清单
├── 文件转换助手.spec         # PyInstaller 打包配置
└── assets/                  # 应用图标
```

## 实现说明

- **PDF 渲染与页面操作**：[PyMuPDF](https://github.com/pymupdf/PyMuPDF)
- **PDF → Word 版式重建**：[pdf2docx](https://github.com/ArtifexSoftware/pdf2docx)（MIT）
- **图片处理**：[Pillow](https://github.com/python-pillow/Pillow)
- **Word → PDF**：通过 `pywin32` 调用本机 Word 导出（`SaveAs FileFormat=17`），转换时会在后台短暂启动 Word 并自动关闭
- 若打包后需要，可在入口处跳过 Qt 析构流程以避免偶发退出崩溃（见 `main.py` 末尾）

## 注意事项

- 转换在本机完成，不依赖任何在线服务
- 加密 PDF 需要密码才能解密或读取
- 批量转换在后台线程执行，界面不会卡死，可随时停止
