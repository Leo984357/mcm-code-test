# Presentation Layout Experiment

**使用 Python 生成可编辑 PowerPoint 页面的小型排版实验。**

当前主要内容是 [main.py](main.py)：以校园咖啡服务为示例，构建一页 16:9 竞争分析幻灯片，包含标题区、竞品对比表和三列要点。仓库名沿用早期实验名称，现有内容聚焦程序化演示文稿生成。

## 文件与产物

| 文件 | 内容 |
|---|---|
| [main.py](main.py) | 通过 `python-pptx` 创建文本框、表格、配色和形状 |
| [diamonds_plot.png](diamonds_plot.png) | 仓库保留的独立图像素材；当前脚本未引用 |

目标产物是可在 PowerPoint 中继续编辑的 `.pptx` 文件。页面中的品牌、业务指标与合作描述是排版示例文案，复用时需替换为有来源的项目内容。

## 使用

需要 Python 3 与 `python-pptx`：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install python-pptx
```

运行前有两处需要调整：

1. 将 `main.py` 末尾的 `prs.save(...)` 改为可写的输出路径，例如 `"competition_analysis.pptx"`。当前默认值是原开发环境的 Windows 桌面路径。
2. 列宽循环使用了 `table.columns[1:]`；`python-pptx` 的列集合仅接受整数索引。将循环改为遍历 `range(1, len(table.columns))`，通过 `table.columns[idx].width` 设置宽度后再运行。

```bash
python main.py
```

也可以导入 `create_competition_analysis_slide(prs)`，向已有 `Presentation` 对象添加页面。生成后在 PowerPoint 中检查字体替换、表格宽度与底部文本框范围；当前脚本按固定坐标排版，长文案需要重新调整位置和字号。
