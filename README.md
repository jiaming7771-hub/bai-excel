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

如果是打包后的软件：

- Windows：结果在 `.exe` 同级的 `output` 文件夹
- macOS：结果在桌面上的 `Excel小工具箱输出` 文件夹

测试文件规模：

- 合并：3 个文件，列顺序不同，其中一个没有“手机号”，合计 1050 行
- 拆分：2200 行，12 个部门
- 清洗：3000+ 行，含重复行、重复手机号和空白行
- 销售：5200 行，20 个销售人员、30 个产品、12 个月

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

### macOS 打包（本机一键）

```bash
./scripts/build_mac.sh
```

生成到桌面 `Excel小工具箱发布/`：

- `Excel小工具箱.app`（可直接双击）
- `Excel小工具箱-macOS.zip`（发给用户）

第一次打开如果提示“无法验证开发者”，请右键选择打开，或到系统设置里允许。

### Windows 打包（必须在 Windows 电脑上执行）

Mac 无法直接打出 Windows 的 `.exe`。把整个 `excel-toolbox` 文件夹拷到 Windows 后：

1. 安装 [Python 3.10+](https://www.python.org/downloads/)（勾选 Add Python to PATH）
2. 双击运行：

```text
scripts\build_windows.bat
```

成功后桌面会出现：

```text
Excel小工具箱发布\Excel小工具箱.exe
```

把这个 `.exe` 发给用户即可，双击就能用，不需要安装 Python。

也可以手动执行：

```bat
pyinstaller --noconfirm --clean excel_toolbox_windows.spec
```

## 说明

- 不依赖 Microsoft Excel
- 不上传任何用户文件
- 没有账号、会员、云端和 AI 功能
- 错误提示使用普通人能看懂的中文
