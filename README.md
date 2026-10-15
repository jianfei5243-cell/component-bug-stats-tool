# 组件 BUG 统计工具

一个离线可视化工具：内置《组件对应模块明细》的“组件 → 模块”字典，上传《1015BUG遗留》后对 M 列“组件”字段进行数量统计，并支持导出 Excel。

## 快速使用

1. 双击打开 `组件BUG统计工具.html`（完全离线，浏览器直接运行）。
2. 组件字典已内置；点击「选择 1015BUG遗留」上传 BUG 表即可统计。
3. 点击右上角“导出 Excel”，生成包含「模块汇总」「组件明细」两个工作表的 `.xlsx`。

> 本仓库不包含任何 BUG 单数据，所有统计结果均由你上传的文件实时计算。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `组件BUG统计工具.html` | 自包含可视化工具（内联 SheetJS + ECharts + 组件字典） |
| `index.html` | 同工具内容，用于 GitHub Pages 在线访问 |
| `build_tool.py` | 生成脚本，可重新生成 HTML（可选导出统计结果） |
| `template.html` | HTML 模板（含占位符，由 `build_tool.py` 填充） |

## 重新构建

```bash
# 仅生成工具（组件字典内置，不含 BUG 数据）
python3 build_tool.py [组件对应模块明细.xlsx]

# 同时基于 BUG 表预导出统计结果
python3 build_tool.py [组件对应模块明细.xlsx] [1015BUG遗留.xlsx]
```

- 默认读取 `~/Downloads/组件对应模块明细.xlsx`。
- 依赖 `openpyxl`；前端库 `xlsx.full.min.js`、`echarts.min.js` 需先下载到 `/tmp/` 供脚本内联。

## 统计口径

- 以组件明细文件的“组件”列为标准字典，对 BUG 表 M 列“组件”字段计数。
- 一条记录关联多个组件（如 `APP-安卓, BSP-NFC`）时，按逗号拆分并分别计入各组件，因此“组件计数合计”可能大于 BUG 记录总数。
- 占比 = 该组件 / 模块的 BUG 数量 ÷ BUG 记录总数。
