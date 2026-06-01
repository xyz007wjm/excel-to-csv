# Excel 转 CSV 工具

将 Excel 文件批量转换为 CSV 格式，分隔符为**分号 (;)**，解决 WPS/Excel 默认导出逗号分隔符的困扰。

## 功能特点

- 支持 .xlsx / .xls 格式
- 批量添加、批量转换
- 分隔符为分号 `;`（不是逗号）
- 输出文件名与原文件同名，放在 `output` 文件夹中
- 编码 UTF-8 BOM（Excel / WPS 打开不乱码）

## 下载

[点此下载最新版](https://github.com/xyz007wjm/excel-to-csv/releases/latest)

| 文件 | 平台 | 说明 |
|------|------|------|
| `ExcelToCSV.exe` | Windows 7/10/11 | 双击运行，无需安装 Python |
| `ExcelToCSV_mac.zip` | macOS | 解压后双击 `Excel转CSV.app` |

## 使用方法

1. 点击 **添加 Excel 文件** 选择要转换的文件
2. 文件列表会显示文件名和工作表数量
3. 点击 **开始转换**
4. 转换后的 CSV 文件生成在源文件目录的 `output` 文件夹中，文件名与原 Excel 文件同名

> 示例：`销售数据.xlsx` → `output/销售数据.csv`

## 打包说明

如需自行打包，克隆仓库后在对应系统执行：

### Windows
```bat
pip install pyinstaller openpyxl xlrd
pyinstaller --onefile --noconsole --name "ExcelToCSV" excel_to_csv.py
```

### macOS
```bash
pip install pyinstaller openpyxl xlrd
pyinstaller --onefile --windowed --name "Excel转CSV" excel_to_csv.py
```

## 软件信息

- 版本：v1.0.1
- 作者：Jack Wang
- 博客：https://wangjinming.com
