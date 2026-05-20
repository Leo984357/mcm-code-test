import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


def create_competition_analysis_slide(prs):
    # 幻灯片尺寸（16:9）
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # 使用空白布局
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # 主题色定义（咖啡主题）
    COLOR_DARK = RGBColor(75, 54, 33)  # 深咖啡（#4B3621）
    COLOR_BEIGE = RGBColor(244, 233, 216)  # 卡其背景（#F4E9D8）
    COLOR_GOLD = RGBColor(212, 175, 55)  # 金铜强调色（#D4AF37）

    # 设置背景
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = COLOR_BEIGE

    # -------------------- 标题区 --------------------
    # 左侧金铜色竖条
    left_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.8), Inches(0.3), Inches(1.2))
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = COLOR_GOLD
    left_bar.line.fill.background()  # 去除边框

    # 主标题
    title = slide.shapes.add_textbox(Inches(1), Inches(0.8), Inches(10), Inches(1))
    tf = title.text_frame
    p = tf.add_paragraph()
    p.text = "竞争分析 —— 构建三维度竞争壁垒"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.alignment = PP_ALIGN.LEFT

    # -------------------- 竞品对比矩阵 --------------------
    # 表格数据
    headers = ["竞争维度", "BrewHive", "校园咖啡馆", "传统售货机", "外卖平台"]
    rows = [
        ["产品力", "现磨+定制化", "现磨+社交", "速溶+标准化", "现磨+丰富选择"],
        ["效率值", "3分钟全流程", "18分钟", "1分钟取货", "30分钟配送"],
        ["场景契合度", "校园全域覆盖", "定点服务", "零散点位", "无物理触点"],
        ["毛利率", "65%", "40%", "50%", "35%"]
    ]

    # 表格位置与尺寸
    table_left = Inches(0.5)
    table_top = Inches(2.2)
    table_width = Inches(12.3)  # 幻灯片宽度-左右边距
    table_height = Inches(2.5)
    cols = len(headers)
    rows_count = len(rows) + 1  # 表头+数据行

    # 添加表格
    table = slide.shapes.add_table(rows_count, cols, table_left, table_top, table_width, table_height).table

    # 填充表头
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_GOLD
        cell.text_frame.paragraphs[0].font.size = Pt(24)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # 填充数据行
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = cell_text
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if (row_idx % 2 == 0) else RGBColor(249, 243,
                                                                                                   233)  # 隔行变色
            cell.text_frame.paragraphs[0].font.size = Pt(22)
            cell.text_frame.paragraphs[0].font.color.rgb = COLOR_DARK
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if col_idx != 0 else PP_ALIGN.LEFT  # 首列左对齐

    # 调整列宽（首列较宽）
    table.columns[0].width = Inches(2.5)
    for col in table.columns[1:]:
        col.width = Inches(2.4)

    # -------------------- 核心壁垒区 --------------------
    # 子标题
    sub_title = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(10), Inches(1))
    tf = sub_title.text_frame
    p = tf.add_paragraph()
    p.text = "核心壁垒构建"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = COLOR_DARK
    p.alignment = PP_ALIGN.LEFT

    # 壁垒内容（分三列布局）
    barriers = [
        {
            "title": "技术壁垒",
            "content": [
                "自研物联网管理系统，设备故障预测准确率92%",
                "AI推荐算法通过10万+杯数据训练，口味匹配度85%"
            ]
        },
        {
            "title": "资源壁垒",
            "content": [
                "与中国高校后勤协会战略合作，优先入驻100所高校",
                "锁定云南小粒咖啡庄园直采，成本较市场价低25%"
            ]
        },
        {
            "title": "模式壁垒",
            "content": [
                "「设备免费投放+流水分成」模式，校方零投入零风险",
                "数据隐私合规：通过ISO 27001认证，消费数据仅用于优化服务"
            ]
        }
    ]

    # 列间距与宽度
    col_width = Inches(3.8)
    col_gap = Inches(0.5)
    start_left = Inches(0.5)
    start_top = Inches(5.8)

    for idx, barrier in enumerate(barriers):
        # 列位置
        left = start_left + idx * (col_width + col_gap)
        top = start_top

        # 标题背景条
        title_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, col_width, Inches(0.6))
        title_bar.fill.solid()
        title_bar.fill.fore_color.rgb = COLOR_GOLD
        title_bar.line.fill.background()

        # 标题文本
        title_text = title_bar.text_frame
        p = title_text.add_paragraph()
        p.text = barrier["title"]
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

        # 内容文本框
        content_box = slide.shapes.add_textbox(left, top + Inches(0.7), col_width, Inches(2))
        tf = content_box.text_frame
        for line in barrier["content"]:
            p = tf.add_paragraph()
            p.text = line
            p.font.size = Pt(20)
            p.font.color.rgb = COLOR_DARK
            p.line_spacing = Pt(28)
            p.space_before = Pt(8)

    return prs


# 使用示例
if __name__ == "__main__":
    prs = Presentation()
    prs.slide_width = Inches(13.33)  # 1920px
    prs.slide_height = Inches(7.5)  # 1080px
    prs = create_competition_analysis_slide(prs)
    prs.save(r"C:\Users\29746\Desktop\竞争分析幻灯片.pptx")
    print("竞争分析幻灯片已生成至桌面")
