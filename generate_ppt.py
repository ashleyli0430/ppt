from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color palette
COLOR_BG_DARK = RGBColor(0x1A, 0x1A, 0x2E)       # Deep navy
COLOR_BG_ACCENT = RGBColor(0x16, 0x21, 0x3E)      # Navy
COLOR_HIGHLIGHT = RGBColor(0xE8, 0x4C, 0x4C)      # Red accent
COLOR_GOLD = RGBColor(0xF5, 0xA6, 0x23)           # Gold
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
COLOR_TEAL = RGBColor(0x0F, 0xB5, 0xA4)           # Teal accent
COLOR_LIGHT_BLUE = RGBColor(0x4A, 0x9E, 0xD4)     # Light blue

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

SLIDE_W = prs.slide_width
SLIDE_H = prs.slide_height

def add_bg(slide, color=COLOR_BG_DARK):
    """Fill slide background."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, color, alpha=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape

def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=COLOR_WHITE,
                 align=PP_ALIGN.LEFT, word_wrap=True):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def add_paragraph(tf, text, font_size=16, bold=False, color=COLOR_WHITE,
                  align=PP_ALIGN.LEFT, space_before=0):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return p

def make_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)

    # Left accent bar
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_HIGHLIGHT)

    # Top decorative strip
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.05), COLOR_GOLD)

    # Center content block
    add_rect(slide, Inches(1.2), Inches(1.8), Inches(10.9), Inches(4.2), COLOR_BG_ACCENT)

    # Title
    tb = slide.shapes.add_textbox(Inches(1.6), Inches(2.1), Inches(10), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "2026年Q1财务策略会议纪要"
    run.font.size = Pt(38)
    run.font.bold = True
    run.font.color.rgb = COLOR_WHITE

    # Subtitle
    add_text_box(slide, "业务激励 · 利润管理 · 组织升级 · IP战略",
                 Inches(1.6), Inches(3.65), Inches(10), Inches(0.6),
                 font_size=20, color=COLOR_GOLD, align=PP_ALIGN.CENTER)

    # Date and participants
    add_text_box(slide, "会议日期：2026年4月1日    参会：刀姐 · AJ · 于老师",
                 Inches(1.6), Inches(4.35), Inches(10), Inches(0.5),
                 font_size=14, color=COLOR_LIGHT_GRAY, align=PP_ALIGN.CENTER)

    # Bottom bar
    add_rect(slide, 0, SLIDE_H - Inches(0.45), SLIDE_W, Inches(0.45), COLOR_BG_ACCENT)
    add_text_box(slide, "DAFA · 内部战略文件  ·  严格保密",
                 Inches(0.3), SLIDE_H - Inches(0.4), Inches(12.5), Inches(0.35),
                 font_size=11, color=COLOR_LIGHT_GRAY, align=PP_ALIGN.CENTER)

def make_agenda_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_TEAL)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "会议议程 AGENDA", Inches(0.3), Inches(0.15),
                 Inches(8), Inches(0.7), font_size=28, bold=True,
                 color=COLOR_WHITE)

    items = [
        ("01", "Q1财务数据回顾", "各业务线收入、利润、成本现状"),
        ("02", "激励机制设计", "利润率标尺 · 爆款奖励 · 团队分润"),
        ("03", "组织转型升级", "个人驱动 → 组织驱动 · 合伙人机制"),
        ("04", "公开课商业模式", "IP费分配 · 导师分润 · 授权体系"),
        ("05", "知识产权规划", "人群战略版权归属 · 商标注册"),
        ("06", "行动项 & 时间线", "本周关键任务与责任人"),
    ]

    cols = 3
    for i, (num, title, desc) in enumerate(items):
        col = i % cols
        row = i // cols
        left = Inches(0.5 + col * 4.25)
        top = Inches(1.3 + row * 2.6)
        w = Inches(3.9)
        h = Inches(2.3)

        add_rect(slide, left, top, w, h, COLOR_BG_ACCENT)
        add_rect(slide, left, top, Inches(0.06), h, COLOR_TEAL)

        # Number
        add_text_box(slide, num, left + Inches(0.15), top + Inches(0.15),
                     Inches(0.7), Inches(0.55), font_size=32, bold=True, color=COLOR_TEAL)
        # Title
        add_text_box(slide, title, left + Inches(0.15), top + Inches(0.75),
                     Inches(3.6), Inches(0.55), font_size=17, bold=True, color=COLOR_WHITE)
        # Desc
        add_text_box(slide, desc, left + Inches(0.15), top + Inches(1.35),
                     Inches(3.6), Inches(0.7), font_size=13, color=COLOR_LIGHT_GRAY)

def make_q1_data_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_GOLD)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "Q1 财务数据回顾", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)
    add_text_box(slide, "2026年1-3月", Inches(10.5), Inches(0.2),
                 Inches(2.5), Inches(0.5), font_size=14, color=COLOR_GOLD,
                 align=PP_ALIGN.RIGHT)

    # Data cards
    cards = [
        ("刀研视频", "~32万", "Q1收入", "净利润率 23%", "月成本 ~5万", COLOR_TEAL, "↑ 健康增长"),
        ("刀姐整体", "226万", "1-2月收入", "同比 +29%", "去年175万", COLOR_GOLD, "↑ 超越去年"),
        ("刀姐视频", "亏损", "当前状态", "成本待配收入", "团队待扩充", COLOR_HIGHLIGHT, "⚠ 需重点关注"),
    ]

    for i, (name, main_num, main_label, stat1, stat2, accent, badge) in enumerate(cards):
        left = Inches(0.4 + i * 4.25)
        top = Inches(1.25)
        w = Inches(4.0)
        h = Inches(4.8)

        add_rect(slide, left, top, w, h, COLOR_BG_ACCENT)
        add_rect(slide, left, top, w, Inches(0.07), accent)

        add_text_box(slide, name, left + Inches(0.2), top + Inches(0.2),
                     Inches(3.5), Inches(0.45), font_size=18, bold=True, color=COLOR_WHITE)

        add_text_box(slide, main_num, left + Inches(0.2), top + Inches(0.8),
                     Inches(3.5), Inches(1.0), font_size=40, bold=True, color=accent)

        add_text_box(slide, main_label, left + Inches(0.2), top + Inches(1.7),
                     Inches(3.5), Inches(0.4), font_size=14, color=COLOR_LIGHT_GRAY)

        # Divider
        add_rect(slide, left + Inches(0.2), top + Inches(2.25), Inches(3.5), Inches(0.02), COLOR_BG_DARK)

        add_text_box(slide, f"• {stat1}", left + Inches(0.2), top + Inches(2.4),
                     Inches(3.5), Inches(0.4), font_size=14, color=COLOR_WHITE)
        add_text_box(slide, f"• {stat2}", left + Inches(0.2), top + Inches(2.9),
                     Inches(3.5), Inches(0.4), font_size=14, color=COLOR_WHITE)

        # Badge
        add_rect(slide, left + Inches(0.2), top + Inches(3.7), Inches(3.0), Inches(0.5), accent)
        add_text_box(slide, badge, left + Inches(0.2), top + Inches(3.72),
                     Inches(3.0), Inches(0.45), font_size=14, bold=True,
                     color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    # Note
    add_rect(slide, Inches(0.4), Inches(6.5), Inches(12.5), Inches(0.5), COLOR_BG_ACCENT)
    add_text_box(slide, "📌  2月份首次实现单月盈利，为近四年来首次 —— 公司财务健康度显著改善",
                 Inches(0.6), Inches(6.52), Inches(12.2), Inches(0.44),
                 font_size=13, color=COLOR_GOLD)

def make_incentive_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_HIGHLIGHT)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "激励机制设计", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)

    # Left: 利润率标尺
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(5.9), Inches(5.8), COLOR_BG_ACCENT)
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(5.9), Inches(0.06), COLOR_HIGHLIGHT)
    add_text_box(slide, "利润率标尺", Inches(0.5), Inches(1.3),
                 Inches(5.5), Inches(0.5), font_size=18, bold=True, color=COLOR_HIGHLIGHT)

    tiers = [
        ("< 15%", "未达基准线，需调整成本或提升收入", COLOR_HIGHLIGHT),
        ("15% - 23%", "基准区间，稳健经营", COLOR_GOLD),
        ("> 23%", "超额利润，触发团队分润机制", COLOR_TEAL),
    ]
    for j, (pct, desc, c) in enumerate(tiers):
        top = Inches(1.95 + j * 1.45)
        add_rect(slide, Inches(0.5), top, Inches(1.6), Inches(1.15), c)
        add_text_box(slide, pct, Inches(0.5), top + Inches(0.3),
                     Inches(1.6), Inches(0.55), font_size=16, bold=True,
                     color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, Inches(2.2), top, Inches(3.7), Inches(1.15), COLOR_BG_DARK)
        add_text_box(slide, desc, Inches(2.35), top + Inches(0.3),
                     Inches(3.4), Inches(0.55), font_size=13, color=COLOR_WHITE)

    # Right: 奖励结构
    add_rect(slide, Inches(6.5), Inches(1.15), Inches(6.5), Inches(5.8), COLOR_BG_ACCENT)
    add_rect(slide, Inches(6.5), Inches(1.15), Inches(6.5), Inches(0.06), COLOR_TEAL)
    add_text_box(slide, "奖励结构设计", Inches(6.7), Inches(1.3),
                 Inches(6.0), Inches(0.5), font_size=18, bold=True, color=COLOR_TEAL)

    reward_items = [
        ("爆款奖金", "单篇突破阈值即时发放\n新档位：6000 → 8000/10000元", COLOR_GOLD),
        ("季度团队奖", "Q1收入1%作为团队池\n按贡献比例分配（编导/剪辑权重可调）", COLOR_TEAL),
        ("超额利润分享", "超过基准利润率的部分\nPD + 核心团队 + 公司 三方分润", COLOR_LIGHT_BLUE),
        ("奖上奖机制", "季度出3-5个爆款\n额外发放绩效奖金（叠加爆款奖之上）", COLOR_HIGHLIGHT),
    ]
    for j, (title, desc, c) in enumerate(reward_items):
        top = Inches(1.95 + j * 1.25)
        add_rect(slide, Inches(6.7), top, Inches(0.06), Inches(1.0), c)
        add_text_box(slide, title, Inches(6.9), top + Inches(0.05),
                     Inches(5.8), Inches(0.4), font_size=14, bold=True, color=c)
        add_text_box(slide, desc, Inches(6.9), top + Inches(0.5),
                     Inches(5.8), Inches(0.45), font_size=12, color=COLOR_LIGHT_GRAY)

def make_org_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_LIGHT_BLUE)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "组织转型升级", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)
    add_text_box(slide, "个人驱动  →  组织驱动", Inches(0.3), Inches(0.6),
                 Inches(9), Inches(0.35), font_size=15, color=COLOR_LIGHT_BLUE)

    # Arrow from old to new
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(5.9), Inches(2.3), COLOR_BG_ACCENT)
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(5.9), Inches(0.06), COLOR_HIGHLIGHT)
    add_text_box(slide, "现状（过去模式）", Inches(0.5), Inches(1.25),
                 Inches(5.5), Inches(0.45), font_size=16, bold=True, color=COLOR_HIGHLIGHT)

    old_items = ["刀姐高度主控，团队执行指令", "各业务线强依赖创始人精力", "激励机制不清晰，团队靠'爱发电'", "缺少操盘者 & HR机制支撑"]
    for j, item in enumerate(old_items):
        add_text_box(slide, f"✗  {item}", Inches(0.5), Inches(1.8 + j * 0.38),
                     Inches(5.5), Inches(0.35), font_size=13, color=COLOR_LIGHT_GRAY)

    add_rect(slide, Inches(6.5), Inches(1.15), Inches(6.5), Inches(2.3), COLOR_BG_ACCENT)
    add_rect(slide, Inches(6.5), Inches(1.15), Inches(6.5), Inches(0.06), COLOR_TEAL)
    add_text_box(slide, "目标（组织驱动）", Inches(6.7), Inches(1.25),
                 Inches(6.0), Inches(0.45), font_size=16, bold=True, color=COLOR_TEAL)

    new_items = ["各业务线独立PD操盘，自主调控成本/收入", "刀法媒体由组织自运转，释放刀姐精力", "利润率标尺 + 分润机制，正向驱动创造", "孵化小AJ，形成多操盘手梯队"]
    for j, item in enumerate(new_items):
        add_text_box(slide, f"✓  {item}", Inches(6.7), Inches(1.8 + j * 0.38),
                     Inches(6.0), Inches(0.35), font_size=13, color=COLOR_WHITE)

    # Roles
    add_text_box(slide, "关键角色定位", Inches(0.3), Inches(3.65),
                 Inches(12.5), Inches(0.45), font_size=17, bold=True, color=COLOR_GOLD)

    roles = [
        ("刀姐", "战略主控 + 刀姐IP业务\n全力dominate并收取IP收益", COLOR_GOLD),
        ("AJ", "潜在合伙人\n整体运营操盘 & 激励体系执行", COLOR_LIGHT_BLUE),
        ("各业务PD", "业务操盘手\n动态调整成本/收入，保利润达标", COLOR_TEAL),
        ("新HR\n(4月15日入职)", "组织建设\n氛围营造 + 机制落地 + 人才激励", COLOR_HIGHLIGHT),
    ]

    for i, (role, desc, c) in enumerate(roles):
        left = Inches(0.3 + i * 3.25)
        top = Inches(4.15)
        w = Inches(3.0)
        h = Inches(2.85)
        add_rect(slide, left, top, w, h, COLOR_BG_ACCENT)
        add_rect(slide, left, top, w, Inches(0.06), c)
        add_text_box(slide, role, left + Inches(0.15), top + Inches(0.2),
                     Inches(2.7), Inches(0.55), font_size=15, bold=True, color=c)
        add_text_box(slide, desc, left + Inches(0.15), top + Inches(0.9),
                     Inches(2.7), Inches(1.7), font_size=12, color=COLOR_WHITE)

def make_open_course_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_GOLD)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "公开课商业模式", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)
    add_text_box(slide, "首期预计 2026年5月上线", Inches(9.5), Inches(0.2),
                 Inches(3.5), Inches(0.5), font_size=14, color=COLOR_GOLD, align=PP_ALIGN.RIGHT)

    # Revenue split diagram
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(12.7), Inches(3.1), COLOR_BG_ACCENT)
    add_text_box(slide, "收入分配模型", Inches(0.5), Inches(1.25),
                 Inches(8), Inches(0.5), font_size=17, bold=True, color=COLOR_GOLD)

    # Pie-like bar
    bar_left = Inches(0.5)
    bar_top = Inches(1.9)
    bar_w = Inches(12.3)
    bar_h = Inches(0.7)
    segments = [
        ("刀姐IP费 10%", 0.10, COLOR_HIGHLIGHT),
        ("刀法品牌IP 10%", 0.10, COLOR_GOLD),
        ("讲师分润 30%", 0.30, COLOR_TEAL),
        ("课程交付成本 20%", 0.20, COLOR_LIGHT_BLUE),
        ("公司留存利润 30%", 0.30, RGBColor(0x5A, 0x5A, 0x8A)),
    ]
    cur = 0
    for label, pct, c in segments:
        seg_w = int(bar_w * pct)
        add_rect(slide, bar_left + cur, bar_top, seg_w, bar_h, c)
        if pct >= 0.09:
            add_text_box(slide, f"{int(pct*100)}%", bar_left + cur + Inches(0.05),
                         bar_top + Inches(0.15), seg_w - Inches(0.1), Inches(0.4),
                         font_size=13, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)
        cur += seg_w

    # Legend
    for i, (label, pct, c) in enumerate(segments):
        col = i % 3
        row = i // 3
        lx = Inches(0.5 + col * 4.1)
        ly = Inches(2.75 + row * 0.38)
        add_rect(slide, lx, ly + Inches(0.05), Inches(0.2), Inches(0.2), c)
        add_text_box(slide, label, lx + Inches(0.3), ly,
                     Inches(3.7), Inches(0.35), font_size=13, color=COLOR_WHITE)

    # Rules
    add_rect(slide, Inches(0.3), Inches(4.4), Inches(6.0), Inches(2.6), COLOR_BG_ACCENT)
    add_rect(slide, Inches(0.3), Inches(4.4), Inches(6.0), Inches(0.06), COLOR_TEAL)
    add_text_box(slide, "讲师合作规则", Inches(0.5), Inches(4.5),
                 Inches(5.6), Inches(0.45), font_size=16, bold=True, color=COLOR_TEAL)

    rules = [
        "外部品牌讲师：约40%（获客+讲课）",
        "主讲师 vs 次讲师按课时比例分配",
        "仅获客不授课：在30%获客池中分配",
        "课件版权另付版权费（版税制度）",
    ]
    for j, r in enumerate(rules):
        add_text_box(slide, f"• {r}", Inches(0.5), Inches(5.1 + j * 0.43),
                     Inches(5.6), Inches(0.4), font_size=13, color=COLOR_WHITE)

    add_rect(slide, Inches(6.6), Inches(4.4), Inches(6.4), Inches(2.6), COLOR_BG_ACCENT)
    add_rect(slide, Inches(6.6), Inches(4.4), Inches(6.4), Inches(0.06), COLOR_GOLD)
    add_text_box(slide, "竞争力参考", Inches(6.8), Inches(4.5),
                 Inches(6.0), Inches(0.45), font_size=16, bold=True, color=COLOR_GOLD)

    comp = [
        "高维学堂：外部讲师40%分成",
        "全职讲师：高底薪 + 年终绩效分成",
        "目标：与高维竞争优质讲师资源",
        "刀法优势：媒体势能 + 精准用户群体",
    ]
    for j, c in enumerate(comp):
        add_text_box(slide, f"• {c}", Inches(6.8), Inches(5.1 + j * 0.43),
                     Inches(6.0), Inches(0.4), font_size=13, color=COLOR_WHITE)

def make_ip_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_TEAL)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "知识产权战略规划", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)

    # IP Structure
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(7.8), Inches(5.8), COLOR_BG_ACCENT)
    add_rect(slide, Inches(0.3), Inches(1.15), Inches(7.8), Inches(0.06), COLOR_TEAL)
    add_text_box(slide, "「人群战略」版权归属架构", Inches(0.5), Inches(1.28),
                 Inches(7.4), Inches(0.5), font_size=17, bold=True, color=COLOR_TEAL)

    # Hierarchy box
    add_rect(slide, Inches(2.5), Inches(2.0), Inches(3.5), Inches(0.75), COLOR_TEAL)
    add_text_box(slide, "科学公司（版权持有方）", Inches(2.5), Inches(2.08),
                 Inches(3.5), Inches(0.55), font_size=14, bold=True,
                 color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    # Lines down
    add_rect(slide, Inches(4.2), Inches(2.75), Inches(0.02), Inches(0.5), COLOR_LIGHT_GRAY)
    add_rect(slide, Inches(1.8), Inches(3.25), Inches(4.8), Inches(0.02), COLOR_LIGHT_GRAY)
    add_rect(slide, Inches(1.8), Inches(3.25), Inches(0.02), Inches(0.4), COLOR_LIGHT_GRAY)
    add_rect(slide, Inches(4.2), Inches(3.25), Inches(0.02), Inches(0.4), COLOR_LIGHT_GRAY)
    add_rect(slide, Inches(6.56), Inches(3.25), Inches(0.02), Inches(0.4), COLOR_LIGHT_GRAY)

    sub_boxes = [
        (Inches(0.8), "刀法咨询\n（授权使用）"),
        (Inches(3.1), "刀法媒体\n（授权使用）"),
        (Inches(5.4), "外部合伙人\n（授权使用）"),
    ]
    for lx, txt in sub_boxes:
        add_rect(slide, lx, Inches(3.65), Inches(1.95), Inches(0.85), COLOR_BG_DARK)
        add_text_box(slide, txt, lx, Inches(3.7),
                     Inches(1.95), Inches(0.75), font_size=12,
                     color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    reg_items = [
        "「人群战略」中文商标（已提交注册）",
        "衍生英文名称及图形商标",
        "核心方法论模型（版权登记）",
        "PPT课件及培训资料（著作权）",
    ]
    add_text_box(slide, "注册范围清单", Inches(0.5), Inches(4.8),
                 Inches(7.4), Inches(0.45), font_size=15, bold=True, color=COLOR_GOLD)
    for j, item in enumerate(reg_items):
        add_text_box(slide, f"□  {item}", Inches(0.5), Inches(5.35 + j * 0.38),
                     Inches(7.4), Inches(0.35), font_size=13, color=COLOR_WHITE)

    # Right: Action plan
    add_rect(slide, Inches(8.4), Inches(1.15), Inches(4.6), Inches(5.8), COLOR_BG_ACCENT)
    add_rect(slide, Inches(8.4), Inches(1.15), Inches(4.6), Inches(0.06), COLOR_GOLD)
    add_text_box(slide, "执行计划", Inches(8.6), Inches(1.28),
                 Inches(4.2), Inches(0.5), font_size=17, bold=True, color=COLOR_GOLD)

    steps = [
        ("责任人", "小凡 + 张经理协助"),
        ("核心资产盘点", "与Paco/天柔确认\n所有IP边界"),
        ("注册提交", "公开课上线前完成\n提交（商标局）"),
        ("内部协议", "科学公司授权各\n主体使用协议"),
        ("对外合作", "外部讲师引用\n人群战略须授权"),
    ]
    for j, (k, v) in enumerate(steps):
        top = Inches(1.9 + j * 1.0)
        add_rect(slide, Inches(8.6), top, Inches(1.3), Inches(0.75), COLOR_GOLD)
        add_text_box(slide, k, Inches(8.6), top + Inches(0.15),
                     Inches(1.3), Inches(0.45), font_size=12, bold=True,
                     color=COLOR_BG_DARK, align=PP_ALIGN.CENTER)
        add_text_box(slide, v, Inches(10.0), top + Inches(0.1),
                     Inches(2.8), Inches(0.55), font_size=12, color=COLOR_WHITE)

def make_action_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, COLOR_BG_DARK)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, COLOR_HIGHLIGHT)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.0), COLOR_BG_ACCENT)

    add_text_box(slide, "行动项 & 时间线", Inches(0.3), Inches(0.15),
                 Inches(9), Inches(0.7), font_size=28, bold=True, color=COLOR_WHITE)

    actions = [
        ("本周四/五", "AJ", "激励方案V1发布", "完整分润机制草案，发至飞书群，三方确认", COLOR_HIGHLIGHT, "高"),
        ("本周内", "云云 + 运营", "Q1爆款奖发放", "调整爆款奖档位(6000→8000/10000)，公开表彰Jasper/琼玉", COLOR_GOLD, "高"),
        ("4月15日", "HR（新入职）", "组织建设启动", "接手团队氛围、绩效沟通、激励机制落地执行", COLOR_TEAL, "中"),
        ("公开课前", "小凡 + 张经理", "IP注册提交", "盘点人群战略全部资产，提交商标及版权注册", COLOR_LIGHT_BLUE, "高"),
        ("本周内", "于老师", "1-2月IP费追溯", "将刀姐IP费从10%追溯调整至最终确定比例", COLOR_GOLD, "中"),
        ("本周内", "AJ + 于老师", "公开课协议框架", "梳理导师分钱比例模板，律师出合同初稿", COLOR_TEAL, "中"),
        ("Q2开始", "AJ + 刀姐", "Q2冲刺计划", "与琼玉/Jasper确认Q2目标与冲刺激励方案", COLOR_LIGHT_BLUE, "中"),
    ]

    for i, (deadline, owner, task, detail, c, priority) in enumerate(actions):
        col = i % 2
        row = i // 2
        if i == 6:
            left = Inches(0.3 + 6.55 * 0)
            top = Inches(1.2 + row * 1.55)
            w = Inches(12.7)
        else:
            left = Inches(0.3 + col * 6.55)
            top = Inches(1.2 + row * 1.55)
            w = Inches(6.2)
        h = Inches(1.35)

        add_rect(slide, left, top, w, h, COLOR_BG_ACCENT)
        add_rect(slide, left, top, Inches(0.06), h, c)

        # Deadline badge
        add_rect(slide, left + Inches(0.15), top + Inches(0.1), Inches(1.1), Inches(0.38), c)
        add_text_box(slide, deadline, left + Inches(0.15), top + Inches(0.1),
                     Inches(1.1), Inches(0.38), font_size=11, bold=True,
                     color=COLOR_BG_DARK, align=PP_ALIGN.CENTER)

        add_text_box(slide, f"[{owner}]  {task}", left + Inches(1.35), top + Inches(0.1),
                     w - Inches(1.6), Inches(0.45), font_size=14, bold=True, color=COLOR_WHITE)
        add_text_box(slide, detail, left + Inches(0.15), top + Inches(0.65),
                     w - Inches(0.3), Inches(0.6), font_size=11, color=COLOR_LIGHT_GRAY)

# Build all slides
make_title_slide(prs)
make_agenda_slide(prs)
make_q1_data_slide(prs)
make_incentive_slide(prs)
make_org_slide(prs)
make_open_course_slide(prs)
make_ip_slide(prs)
make_action_slide(prs)

output = "/home/user/ppt/Q1财务策略会议纪要.pptx"
prs.save(output)
print(f"Saved: {output}")
