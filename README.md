# 组件 BUG 统计工具

一个离线可视化工具：依据《组件对应模块明细》的“组件”清单，对《1015BUG遗留》M 列“组件”字段进行数量统计，并支持导出 Excel。

## 快速使用

1. 双击打开 `组件BUG统计工具.html`（完全离线，浏览器直接运行）。
2. 页面默认已内置示例数据；也可以点击上传新的组件明细表与 BUG 表实时重算。
3. 点击右上角“导出 Excel”，生成包含「模块汇总」「组件明细」两个工作表的 `.xlsx`。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `组件BUG统计工具.html` | 自包含可视化工具（内联 SheetJS + ECharts + 示例数据） |
| `组件BUG统计_结果.xlsx` | 基于内置示例数据预导出的统计结果 |
| `build_tool.py` | 生成脚本，可重新统计并生成 HTML / Excel |
| `template.html` | HTML 模板（含占位符，由 `build_tool.py` 填充） |

## 重新构建

```bash
python3 build_tool.py [组件对应模块明细.xlsx] [1015BUG遗留.xlsx]
```

- 默认读取 `~/Downloads/组件对应模块明细.xlsx` 与 `~/Downloads/1015BUG遗留.xlsx`。
- 依赖 `openpyxl`（生成统计与 Excel）。
- 前端库 `xlsx.full.min.js`、`echarts.min.js` 需先下载到 `/tmp/` 供脚本内联（脚本中 `XLSX_LIB` / `ECHARTS_LIB` 路径可改）。

## 统计口径

- 以组件明细文件的“组件”列为标准字典，对 BUG 表 M 列“组件”字段计数。
- M 列中如 `APP-安卓, BSP-NFC` 这类一条记录关联多个组件时，按逗号拆分并分别计入各组件，因此“组件计数合计”可能大于 BUG 记录总数。
- 占比 = 该组件 / 模块的 BUG 数量 ÷ BUG 记录总数。
