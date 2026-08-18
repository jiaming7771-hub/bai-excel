# Excel小工具箱

批量处理 Excel，让重复工作一键完成。

这是给普通办公室用户用的本地小工具，不需要会编程，也不需要安装 Microsoft Excel。全程在电脑本地处理，不联网、不上传文件。

## 能做什么

1. Excel 批量合并：多个表按列名合并成一个总表
2. Excel 数据拆分：按部门等字段拆成多个 Excel
3. 重复数据清洗：删除重复行、空白行，或按手机号等字段去重
4. 销售数据汇总：自动统计销售额、数量、人员排名和月度销售

## 运行环境

- Windows 10/11 或 macOS
- 开发时需要 Python 3.10+
- 用户使用打包后的软件时，不需要安装 Python

## 开发运行

```bash
cd excel-toolbox
python3 -m venv .venv
# Windows:
# .venv\Scripts\activate
# macOS:
source .venv/bin/activate

pip install -r requirements.txt
python main.py
```

## 生成测试数据并测试

```bash
python testdata/generate_testdata.py
python -m pytest -q
```

测试数据会生成在 `testdata/`：

- `testdata/test_merge/sales_a.xlsx`
- `testdata/test_merge/sales_b.xlsx`
- `testdata/test_merge/sales_c.xlsx`
- `testdata/test_split.xlsx`
- `testdata/test_clean.xlsx`
- `testdata/test_sales.xlsx`

处理结果默认保存在 `output/`。

## 给普通用户怎么用

1. 打开「Excel小工具箱」
2. 点首页四个功能里的「立即使用」
3. 选择文件，或把 Excel 拖进窗口
4. 点开始按钮
5. 完成后点「打开输出文件夹」

## 打包（用户不需要安装 Python）

先安装依赖和 PyInstaller：

```bash
pip install -r requirements.txt
```

### Windows 打包命令

在 Windows 电脑上执行：

```bat
pyinstaller --noconfirm --clean --windowed --onefile --name "Excel小工具箱" --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd main.py
```

生成文件：

```text
dist\Excel小工具箱.exe
```

也可以使用仓库里的 spec：

```bat
pyinstaller --noconfirm excel_toolbox.spec
```

### macOS 打包命令

在 Mac 上执行：

```bash
pyinstaller --noconfirm --clean --windowed --name "Excel小工具箱" --hidden-import pandas --hidden-import openpyxl --hidden-import xlrd main.py
```

生成文件：

```text
dist/Excel小工具箱.app
```

如果需要单个可执行文件，也可以加 `--onefile`，但 macOS 更推荐 `.app` 文件夹形式。

第一次打开如果提示“无法验证开发者”，请在系统设置里允许打开，或右键选择打开。

## 说明

- 不依赖 Microsoft Excel
- 不上传任何用户文件
- 没有账号、会员、云端和 AI 功能
- 错误提示使用普通人能看懂的中文
