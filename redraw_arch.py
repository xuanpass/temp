"""
重绘 RZY 靶场平台架构图 v2
- 删除原 OLE/Visio 对象
- 用 python-pptx 原生形状重绘，保持可编辑
"""
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# === 配色方案 ===
C_BG_CARD = RGBColor(0x0E, 0x2A, 0x4A)
C_BORDER_CONTAINER = RGBColor(0x2E, 0x75, 0xB6)
C_BORDER_CARD = RGBColor(0x3A, 0x9B, 0xD9)
C_ACCENT = RGBColor(0x00, 0xCC, 0xFF)
C_TITLE = RGBColor(0xA2, 0xFF, 0xFF)
C_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
C_TEXT_DIM = RGBColor(0x8C, 0xB8, 0xD8)

# === 加载原 PPT ===
ppt_path = ".uploads/121267fd-b4bc-46d4-8e17-df7133ae1cfa_【10.21】RZY靶场训练项目介绍V4.2 - 副本 (2).pptx"
prs = Presentation(ppt_path)
slide = prs.slides[0]
SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height
print(f"Slide: {SLIDE_W} x {SLIDE_H}")

# === 第1步：删除旧形状 ===
spTree = slide.shapes._spTree

gf_nodes = spTree.findall(qn('p:graphicFrame'))
print(f"Removing {len(gf_nodes)} graphicFrame(s)")
for gf in gf_nodes:
    spTree.remove(gf)

to_remove = []
for sp in spTree.findall(qn('p:sp')):
    cNvPr = sp.find('.//' + qn('p:cNvPr'))
    if cNvPr is not None:
        name = cNvPr.get('name', '')
        if name in ('TextBox 3', 'AutoShape 19'):
            to_remove.append(sp)
            print(f"Removing shape: {name}")
for node in to_remove:
    spTree.remove(node)

prs.save("/workspace/rzy_temp.pptx")
prs = Presentation("/workspace/rzy_temp.pptx")
slide = prs.slides[0]

# === 布局 ===
# 16:9 slide ~ 13.33" x 7.5"
# 留够边距避免越界
MARGIN_L = Inches(1.1)       # 左侧 + 层标签
MARGIN_R = Inches(0.45)      # 右侧 + 侧栏标签
CONTENT_X = MARGIN_L
CONTENT_W = SLIDE_W - MARGIN_L - MARGIN_R

TOP_Y = Inches(1.65)
BOTTOM_Y = Inches(7.10)
TOTAL_H = BOTTOM_Y - TOP_Y

LAYERS = 4
GAP_H = Inches(0.12)
LAYER_H = (TOTAL_H - GAP_H * (LAYERS - 1)) // LAYERS

# === 辅助函数 ===
def add_rounded_rect(slide, x, y, w, h, fill_color=None, border_color=None, border_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.shadow.inherit = False
    shape.shadow.visible = False
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width or Pt(0.75)
    else:
        shape.line.fill.background()
    # 去掉圆角？保持圆角
    return shape

def add_text(slide, x, y, w, h, text, font_size=10, bold=False, color=C_TEXT, align=PP_ALIGN.CENTER, font_name="微软雅黑"):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    txBox.line.fill.background()
    txBox.fill.background()
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')
    return txBox

def add_text_vt(slide, x, y, w, h, text, font_size=11, bold=True, color=C_ACCENT):
    """竖排文字"""
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    txBox.line.fill.background()
    txBox.fill.background()
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')
        bodyPr.set('vert', 'eaVert')
    return txBox

def add_card(slide, x, y, w, h, text, font_size=9, fill=C_BG_CARD, border=C_BORDER_CARD, text_color=C_TEXT, bold=False):
    add_rounded_rect(slide, x, y, w, h, fill_color=fill, border_color=border, border_width=Pt(0.75))
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = text_color
    run.font.name = "微软雅黑"
    txBox.line.fill.background()
    txBox.fill.background()
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('anchor', 'ctr')
    return txBox

# === 颜色渐变 ===
layer_colors = [
    {"fill": RGBColor(0x0D, 0x3B, 0x6B), "border": C_ACCENT},
    {"fill": RGBColor(0x0F, 0x33, 0x58), "border": C_BORDER_CONTAINER},
    {"fill": RGBColor(0x11, 0x2D, 0x4C), "border": C_BORDER_CONTAINER},
    {"fill": RGBColor(0x13, 0x28, 0x42), "border": C_BORDER_CONTAINER},
]

# === 绘制每层 ===
def draw_layer(slide, idx, name, sub_label_top=""):
    y = TOP_Y + idx * (LAYER_H + GAP_H)
    x = CONTENT_X
    w = CONTENT_W
    h = LAYER_H
    colors = layer_colors[idx]

    # 容器
    add_rounded_rect(slide, x, y, w, h,
                     fill_color=colors["fill"],
                     border_color=colors["border"],
                     border_width=Pt(1))

    # 层标签（竖排）
    label_w = Inches(0.5)
    label_x = x + Inches(0.12)
    add_text_vt(slide, label_x, y + Inches(0.05), label_w, h - Inches(0.1),
                name, font_size=10, bold=True, color=C_ACCENT)

    # 模块区
    mx = label_x + label_w + Inches(0.08)
    mw = w - label_w - Inches(0.2)
    my = y + Inches(0.08)
    mh = h - Inches(0.16)

    return mx, my, mw, mh


# === 大标题 ===
add_text(slide, Inches(0.4), Inches(0.25), Inches(5), Inches(0.8),
         "系统架构", font_size=28, bold=True, color=C_ACCENT, align=PP_ALIGN.LEFT)
add_text(slide, Inches(0.4), Inches(0.95), Inches(8), Inches(0.6),
         "RZY 靶场平台架构", font_size=22, bold=True, color=C_TITLE, align=PP_ALIGN.LEFT)

# ====== 层 0：服务应用层 ======
mx, my, mw, mh = draw_layer(slide, 0, "服务应用层")
modules = ["理论训练", "模拟演练", "红蓝对抗", "复盘评估"]
n = len(modules)
gap = Inches(0.12)
cw = (mw - gap * (n - 1)) / n
ch = mh
for j, m in enumerate(modules):
    cx = mx + j * (cw + gap)
    add_card(slide, cx, my, cw, ch, m, font_size=12, bold=True, border=C_ACCENT)

# ====== 层 1：功能支撑层 ======
mx, my, mw, mh = draw_layer(slide, 1, "功能支撑层")

# 分成 4 行：row1 / row2 / sub_label(中间标签) / row3
row_gap = Inches(0.06)

# row1
row1_y = my
row1_h = mh * 0.27 - row_gap / 2
row1 = ["场景构建", "导调控制", "事件模拟", "态势展示", "效能模拟"]
n = len(row1)
gap = Inches(0.07)
cw = (mw - gap * (n - 1)) / n
for j, m in enumerate(row1):
    cx = mx + j * (cw + gap)
    add_card(slide, cx, row1_y, cw, row1_h, m, font_size=10, bold=True)

# row2
row2_y = row1_y + row1_h + row_gap
row2_h = mh * 0.27 - row_gap / 2
row2 = ["战术自动化支撑", "数据采集支撑", "素材管理", "对抗管理", "系统管理"]
for j, m in enumerate(row2):
    cx = mx + j * (cw + gap)
    add_card(slide, cx, row2_y, cw, row2_h, m, font_size=10)

# sub_label（中间标签，放在 row2 和 row3 之间，居中）
row3_y = row2_y + row2_h + row_gap
row3_h = mh - row1_h - row2_h - row_gap * 2

# 中间标签 "平台管理功能模块" 带箭头 ↑，放在 row3 卡片的正上方
mid_label_w = Inches(2.4)
mid_label_x = mx + (mw - mid_label_w) / 2
mid_label_y = row3_y - Inches(0.18)  # 在 row3 上方一点
add_text(slide, mid_label_x, mid_label_y, mid_label_w, Inches(0.18),
         "↑ 平台管理功能模块 ↑", font_size=8, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)

# row3（三个宽卡片）
row3 = ["多轮传播推演引擎", "推演数据库", "认知演练可视化"]
n3 = len(row3)
gap3 = Inches(0.1)
cw3 = (mw - gap3 * (n3 - 1)) / n3
for j, m in enumerate(row3):
    cx = mx + j * (cw3 + gap3)
    add_card(slide, cx, row3_y, cw3, row3_h, m, font_size=10, bold=True)

# ====== 层 2：社交智能体支撑层 ======
mx, my, mw, mh = draw_layer(slide, 2, "社交智能体支撑层")

row_gap = Inches(0.06)

# sub_label 顶部标签
top_label_y = my
top_label_h = Inches(0.18)
add_text(slide, mx, top_label_y, Inches(2.5), top_label_h,
         "智能体指令体系", font_size=9, bold=True, color=C_ACCENT, align=PP_ALIGN.LEFT)

# row1（智能体）
row1_y = top_label_y + top_label_h + Inches(0.04)
row1_h = mh * 0.28 - row_gap / 2 - top_label_h - Inches(0.04)
agents = ["平台智能体", "个体智能体", "群体智能体", "导调智能体", "攻防智能体"]
n = len(agents)
gap = Inches(0.09)
cw = (mw - gap * (n - 1)) / n
for j, m in enumerate(agents):
    cx = mx + j * (cw + gap)
    add_card(slide, cx, row1_y, cw, row1_h, m, font_size=10, bold=True)

# row2（7个互动）
row2_y = row1_y + row1_h + row_gap
row2_h = mh * 0.26
interactions = ["制品投送", "浏览互动", "点赞互动", "评论互动", "分享互动", "加关注", "加好友"]
n_int = len(interactions)
int_gap = Inches(0.05)
int_w = (mw - int_gap * (n_int - 1)) / n_int
for j, m in enumerate(interactions):
    cx = mx + j * (int_w + int_gap)
    add_card(slide, cx, row2_y, int_w, row2_h, m, font_size=9)

# row3（4个模拟）
row3_y = row2_y + row2_h + row_gap
row3_h = my + mh - row3_y
sims = ["推荐模拟", "封禁模拟", "转发模拟", "观点转变"]
n_sims = len(sims)
sim_gap = Inches(0.12)
sim_w = (mw - sim_gap * (n_sims - 1)) / n_sims
for j, m in enumerate(sims):
    cx = mx + j * (sim_w + sim_gap)
    add_card(slide, cx, row3_y, sim_w, row3_h, m, font_size=10)

# ====== 层 3：资源支撑层 ======
mx, my, mw, mh = draw_layer(slide, 3, "资源支撑层")

row_gap = Inches(0.05)

# row1：2个大卡片
row1_y = my
big_h = mh * 0.34 - row_gap / 2
big_gap = Inches(0.1)
big_w = (mw - big_gap) / 2
add_card(slide, mx, row1_y, big_w, big_h,
         "统一数据资源编排管理", font_size=10, fill=C_BG_CARD, border=C_ACCENT, bold=True)
add_card(slide, mx + big_w + big_gap, row1_y, big_w, big_h,
         "虚拟化仿真平台构建", font_size=10, fill=C_BG_CARD, border=C_ACCENT, bold=True)

# row2：左4个 + 右6个
row2_y = row1_y + big_h + row_gap
row_h2 = mh * 0.28

left_w = mw * 0.36
left_gap = Inches(0.07)
left_modules = ["平台资源", "策略资源", "事件数据", "账号数据"]
n_l = len(left_modules)
lw = (left_w - left_gap * (n_l - 1)) / n_l
for j, m in enumerate(left_modules):
    cx = mx + j * (lw + left_gap)
    add_card(slide, cx, row2_y, lw, row_h2, m, font_size=9)

right_x = mx + left_w + Inches(0.12)
right_w = mx + mw - right_x
logos = ["FB", "Youtube", "Tiktok", "X", "抖音", "Linkedin"]
n_logo = len(logos)
logo_gap = Inches(0.06)
logo_w = (right_w - logo_gap * (n_logo - 1)) / n_logo
for j, logo in enumerate(logos):
    cx = right_x + j * (logo_w + logo_gap)
    add_card(slide, cx, row2_y, logo_w, row_h2, logo, font_size=9,
             fill=RGBColor(0x0A, 0x24, 0x40), border=C_ACCENT)

# row3：左2个 + 右6个
row3_y = row2_y + row_h2 + row_gap
row_h3 = my + mh - row3_y

left2 = ["实网数据获取", "数据标签化治理"]
n_l2 = len(left2)
lw2 = (left_w - left_gap * (n_l2 - 1)) / n_l2
for j, m in enumerate(left2):
    cx = mx + j * (lw2 + left_gap)
    add_card(slide, cx, row3_y, lw2, row_h3, m, font_size=9)

logos2 = ["京东", "淘宝", "邮政", "PTT", "LD", "GD"]
for j, logo in enumerate(logos2):
    cx = right_x + j * (logo_w + logo_gap)
    add_card(slide, cx, row3_y, logo_w, row_h3, logo, font_size=9,
             fill=RGBColor(0x0A, 0x24, 0x40), border=C_ACCENT)

# === 右侧竖排标签 ===
right_x = SLIDE_W - Inches(0.4)  # 从右边留 0.4"
right_w = Inches(0.35)
half_h = (BOTTOM_Y - TOP_Y) / 2

add_text_vt(slide, right_x, TOP_Y, right_w, half_h,
            "标准体系", font_size=10, bold=True, color=C_ACCENT)
add_text_vt(slide, right_x, TOP_Y + half_h, right_w, half_h,
            "安全保密体系", font_size=10, bold=True, color=C_ACCENT)

# === 保存 ===
output_path = "/workspace/RZY靶场平台架构_美化版.pptx"
prs.save(output_path)
print(f"\nSaved to: {output_path}")
