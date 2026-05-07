from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

# ── Palette ───────────────────────────────────────────────────
C_RED    = RGBColor(0xC0, 0x00, 0x00)
C_DARK   = RGBColor(0x10, 0x14, 0x24)
C_NAVY   = RGBColor(0x16, 0x21, 0x3E)
C_BLUE   = RGBColor(0x0F, 0x3C, 0x78)
C_ACCENT = RGBColor(0x00, 0xB4, 0xD8)
C_GOLD   = RGBColor(0xF5, 0xA6, 0x23)
C_GREEN  = RGBColor(0x2D, 0xC6, 0x53)
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xB0, 0xC4, 0xDE)
C_MGRAY  = RGBColor(0x8A, 0x9B, 0xB0)
C_DGRAY  = RGBColor(0x44, 0x55, 0x66)

BLANK = prs.slide_layouts[6]

# ── Helpers ───────────────────────────────────────────────────
def rect(slide, x, y, w, h, fill=None, line=None):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line
    else:
        s.line.fill.background()
    return s

def txt(slide, text, x, y, w, h, size=14, bold=False, color=C_WHITE,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tb.word_wrap = wrap
    tf = tb.text_frame; tf.word_wrap = wrap
    p  = tf.paragraphs[0]; p.alignment = align
    r  = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold
    r.font.color.rgb = color; r.font.italic = italic
    return tb

def bg(slide):
    rect(slide, 0, 0, 13.33, 7.5, fill=C_DARK)

def top_bar(slide, title, subtitle="", tag_text=""):
    rect(slide, 0, 0, 13.33, 1.05, fill=C_NAVY)
    rect(slide, 0, 0, 0.2, 1.05, fill=C_RED)
    if tag_text:
        rect(slide, 0.35, 0.15, len(tag_text)*0.13+0.3, 0.3, fill=C_RED)
        txt(slide, tag_text, 0.45, 0.17, len(tag_text)*0.13+0.2, 0.26,
            size=10, bold=True)
        txt(slide, title, 0.35, 0.52, 12.5, 0.55, size=24, bold=True)
    else:
        txt(slide, title, 0.35, 0.22, 12.5, 0.55, size=26, bold=True)
    if subtitle:
        txt(slide, subtitle, 0.35, 0.72, 10, 0.32, size=13, color=C_LGRAY)

def bottom_bar(slide, left="遠傳 DTIS｜AI 技術學院", right=""):
    rect(slide, 0, 6.9, 13.33, 0.6, fill=C_RED)
    txt(slide, left,  0.2,  6.93, 6, 0.45, size=11, bold=True)
    if right:
        txt(slide, right, 10.5, 6.93, 2.7, 0.45, size=10, align=PP_ALIGN.RIGHT)

def hline(slide, x, y, w, color=C_RED, h=0.045):
    rect(slide, x, y, w, h, fill=color)

def tag(slide, text, x, y, bg_c=C_RED, fg=C_WHITE, size=10):
    w = len(text)*0.115 + 0.32
    rect(slide, x, y, w, 0.3, fill=bg_c)
    txt(slide, text, x+0.1, y+0.03, w-0.12, 0.26, size=size, bold=True, color=fg)

def stat_card(slide, val, label, desc, x, y, w, h, col):
    rect(slide, x, y, w, h, fill=C_NAVY)
    rect(slide, x, y, w, 0.07, fill=col)
    txt(slide, val,   x+0.2, y+0.1,  w-0.3, 0.85, size=44, bold=True, color=col)
    txt(slide, label, x+0.2, y+0.95, w-0.3, 0.38, size=14, bold=True, color=C_WHITE)
    txt(slide, desc,  x+0.2, y+1.35, w-0.3, 0.5,  size=11, color=C_LGRAY)

def case_card(slide, num, title, pain, solution, result, x, y, w, h, col):
    rect(slide, x, y, w, h, fill=C_NAVY)
    rect(slide, x, y, 0.1, h, fill=col)
    rect(slide, x, y, w, 0.06, fill=col)
    txt(slide, f"案例 {num}", x+0.2, y+0.08, w, 0.28, size=10, bold=True, color=col)
    txt(slide, title,        x+0.2, y+0.38, w-0.3, 0.45, size=16, bold=True)
    hline(slide, x+0.2, y+0.88, w-0.4, color=col)
    tag(slide, "痛點", x+0.2, y+1.0, bg_c=RGBColor(0x7F,0x1D,0x1D))
    txt(slide, pain,     x+0.2, y+1.38, w-0.3, 0.75, size=11.5, color=C_LGRAY)
    tag(slide, "解法", x+0.2, y+2.2, bg_c=C_BLUE)
    txt(slide, solution, x+0.2, y+2.58, w-0.3, 0.75, size=11.5, color=C_WHITE)
    tag(slide, "成果", x+0.2, y+3.4, bg_c=RGBColor(0x14,0x53,0x29))
    txt(slide, result,   x+0.2, y+3.78, w-0.3, 0.75, size=11.5, color=C_GREEN)


# ══════════════════════════════════════════════════════════════
# Slide 1 ── Cover
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
rect(s, 0, 0, 0.2, 7.5, fill=C_RED)
rect(s, 0, 0, 13.33, 0.06, fill=C_RED)
rect(s, 0.2, 0, 13.13, 7.5, fill=C_DARK)

# Grid lines (decorative)
for gx in [3,6,9,12]:
    rect(s, gx, 0, 0.012, 7.5, fill=RGBColor(0x22,0x2D,0x45))
for gy in [1.5,3,4.5,6]:
    rect(s, 0.2, gy, 13.13, 0.012, fill=RGBColor(0x22,0x2D,0x45))

rect(s, 0.5, 1.6, 7.5, 4.5, fill=RGBColor(0x12,0x18,0x2E))
rect(s, 0.5, 1.6, 0.15, 4.5, fill=C_RED)

tag(s, "DTIS AI 技術學院", 0.85, 1.75, bg_c=C_RED, size=12)
txt(s, "AI 驅動工程未來", 0.85, 2.25, 7, 1.0, size=44, bold=True)
txt(s, "DTIS 已在路上", 0.85, 3.25, 7, 1.0, size=44, bold=True, color=C_GOLD)
hline(s, 0.85, 4.38, 6.5)
txt(s, "「我們不是在導入 AI 工具，我們是在重新定義\n 軟體開發的作業系統。」",
    0.85, 4.52, 7.0, 1.1, size=14, italic=True, color=C_LGRAY)
txt(s, "轉型成果報告 ｜ 2026 Q2", 0.85, 5.75, 5, 0.4, size=13, color=C_MGRAY)

# Right stat preview
for i,(v,l,c) in enumerate([("105","人 AI 社群",C_ACCENT),
                              ("20+","項 Skills",C_GOLD),
                              ("100%","自動化覆蓋",C_GREEN)]):
    rx=8.5; ry=1.9+i*1.65
    rect(s, rx, ry, 4.5, 1.45, fill=C_NAVY)
    rect(s, rx, ry, 0.1, 1.45, fill=c)
    txt(s, v, rx+0.3, ry+0.1, 2.5, 0.9, size=42, bold=True, color=c)
    txt(s, l, rx+0.3, ry+0.98, 3.8, 0.38, size=14, color=C_WHITE)

bottom_bar(s, right="內部機密")


# ══════════════════════════════════════════════════════════════
# Slide 2 ── Executive Summary
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "一年內，從零到規模",
        "社群從 12 人擴展到 105 人，AI 能力已系統化、平台化、可複製",
        "執行摘要")

stats = [
    ("12 → 105", "AI 社群規模", "三個月成長 8.75 倍", C_ACCENT),
    ("20+",       "Skills 封裝", "可重用 AI 能力模組",  C_GOLD),
    ("10+",       "MCP Services","上線運行中",          C_RED),
    ("5",         "SDLC 環節",   "全程 AI Agent 覆蓋", C_GREEN),
    ("100%",      "自動化覆蓋",  "AEM 全站測試達成",   RGBColor(0xA0,0x55,0xFF)),
]
for i,(v,l,d,c) in enumerate(stats):
    col = i % 3; row = i // 3
    sx = 0.45 + col*4.3; sy = 1.35 + row*2.55
    stat_card(s, v, l, d, sx, sy, 4.0, 2.25, c)

# SCQA bottom strip
rect(s, 0.4, 6.3, 12.5, 0.52, fill=C_BLUE)
txt(s, "挑戰：SDLC 每個環節依賴大量人工，交付週期長、風險難以前置化管理。"
    "  →  答案：系統性改造 SDLC，三個月落地四項可量化 AI 成果",
    0.6, 6.35, 12.1, 0.42, size=12.5, bold=True, color=C_ACCENT)
bottom_bar(s, right="2 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 3 ── AI Collaboration Architecture
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "不是工具堆疊，是有機協作系統",
        "LLM 是大腦、Skills 是知識、MCP 是手腳，三層架構讓 AI 真正會做事", "核心架構")

layers = [
    ("LLM  大腦", "思考 / 推理 / 判斷 / 決策",
     "Claude / GPT / Gemini\n負責「想清楚」", C_ACCENT, 1.3),
    ("Skills  專業知識", "封裝工程師 Domain Know-how",
     "Requirement Auditor / Code Reviewer\n負責「知道什麼」", C_GOLD, 2.85),
    ("MCP  手腳", "標準化工具協定 — AI 世界的 HTTP 協議",
     "Playwright / PostgreSQL / 維運 API\n負責「動起來做」", C_RED, 4.4),
]
for (title, sub, body, col, ly) in layers:
    rect(s, 0.6, ly, 7.5, 1.35, fill=C_NAVY)
    rect(s, 0.6, ly, 0.12, 1.35, fill=col)
    rect(s, 0.6, ly, 7.5, 0.06, fill=col)
    txt(s, title, 0.9, ly+0.1, 4.5, 0.45, size=18, bold=True, color=col)
    txt(s, sub,   0.9, ly+0.55, 7.0, 0.35, size=12, color=C_LGRAY)
    txt(s, body,  0.9, ly+0.88, 7.0, 0.45, size=11.5, color=C_WHITE)

# Arrows
for ay in [2.68, 4.23]:
    txt(s, "▼", 3.85, ay, 0.5, 0.26, size=16, color=C_RED, align=PP_ALIGN.CENTER)

# Right panel
rect(s, 8.55, 1.25, 4.5, 5.35, fill=C_NAVY)
rect(s, 8.55, 1.25, 4.5, 0.06, fill=C_ACCENT)
txt(s, "三層分工的戰略價值", 8.75, 1.35, 4.2, 0.45, size=15, bold=True, color=C_ACCENT)
points = [
    ("可維護", "每層獨立升級，不影響其他層"),
    ("可擴充", "新增 Skills / MCP 不改架構"),
    ("可審計", "每次工具呼叫皆有紀錄"),
    ("不鎖廠商", "LLM 可替換，技術自主"),
    ("知識傳承", "Skills 封裝資深工程師經驗"),
]
for i,(pt,desc) in enumerate(points):
    py = 1.9 + i*0.9
    tag(s, pt, 8.75, py, bg_c=C_BLUE, size=10)
    txt(s, desc, 8.75, py+0.35, 4.1, 0.38, size=11.5, color=C_LGRAY)

bottom_bar(s, right="3 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 4 ── SDLC
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "從需求到維運，AI 全程在場",
        "每個工程環節都有專屬 AI Agent，沒有人力瓶頸，沒有知識孤島", "SDLC AI 化")

stages = [
    ("需求", "Requirement\nAI Auditor",
     "拆解需求\n識別模糊點\n生成 User Story", C_GOLD),
    ("開發", "AI Pair\nProgramming",
     "程式碼生成\n架構建議\n重構輔助", C_ACCENT),
    ("上版", "Security &\nLogic Auditor",
     "PR 雙層審核\n風險分級報告\nHigh Risk 直接退回", C_RED),
    ("測試", "Auto Testing\nAgent",
     "UI 自動回歸\nPlaywright+MCP\n持續覆蓋", C_GREEN),
    ("維運", "Diagnostic &\nAnalysis Agent",
     "主動預警\n多 MCP 推論\n知識封裝傳承", RGBColor(0xA0,0x55,0xFF)),
]
for i,(stage,agent,body,col) in enumerate(stages):
    sx = 0.35 + i*2.57
    # Stage box
    rect(s, sx, 1.15, 2.35, 0.7, fill=col)
    txt(s, stage, sx, 1.15, 2.35, 0.7, size=20, bold=True, align=PP_ALIGN.CENTER)
    # Arrow
    if i < 4:
        txt(s, "→", sx+2.35, 1.32, 0.3, 0.38, size=16, color=C_MGRAY, align=PP_ALIGN.CENTER)
    # Agent name box
    rect(s, sx, 2.0, 2.35, 1.05, fill=C_NAVY)
    rect(s, sx, 2.0, 2.35, 0.06, fill=col)
    txt(s, agent, sx+0.1, 2.1, 2.2, 0.95, size=12.5, bold=True, color=col, align=PP_ALIGN.CENTER)
    # Detail box
    rect(s, sx, 3.18, 2.35, 2.1, fill=RGBColor(0x12,0x18,0x2E))
    rect(s, sx, 3.18, 0.08, 2.1, fill=col)
    txt(s, body, sx+0.18, 3.28, 2.1, 1.95, size=11.5, color=C_WHITE)

# MCPHub bottom
rect(s, 0.35, 5.45, 12.65, 0.72, fill=C_BLUE)
txt(s, "MCPHub 統一工具平台", 0.55, 5.5, 3.5, 0.38, size=13, bold=True, color=C_ACCENT)
txt(s, "所有 Agent 共享 10+ MCP Services ｜ 智慧路由 ｜ 安全認證 ｜ 即時監控",
    4.2, 5.55, 8.6, 0.55, size=12.5, color=C_WHITE)

bottom_bar(s, right="4 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 5 ── MCPHub
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "MCPHub：AI 能力的統一接線盒",
        "讓所有 Agent 共享工具、共享知識，是規模化的關鍵基礎", "技術平台")

# Hub circle (simulated with square)
rect(s, 4.65, 2.3, 4.0, 3.2, fill=C_NAVY)
rect(s, 4.65, 2.3, 4.0, 0.08, fill=C_ACCENT)
txt(s, "MCPHub", 4.65, 2.42, 4.0, 0.65, size=26, bold=True, color=C_ACCENT, align=PP_ALIGN.CENTER)
txt(s, "AI Agent 世界的\nAPI Gateway + Control Plane",
    4.65, 3.12, 4.0, 0.85, size=13, color=C_WHITE, align=PP_ALIGN.CENTER)
hline(s, 5.0, 4.05, 3.3, color=C_ACCENT)
for i,(feat,val) in enumerate([("MCP Services","10+"),("Agent 共用","全部"),("監控告警","即時")]):
    fx = 5.05 + i*1.25
    txt(s, val,  fx, 4.15, 1.15, 0.42, size=18, bold=True, color=C_GOLD, align=PP_ALIGN.CENTER)
    txt(s, feat, fx, 4.55, 1.15, 0.3,  size=9,  color=C_LGRAY, align=PP_ALIGN.CENTER)

# Left services
left_svcs = [("Playwright MCP","UI 自動化測試",C_GREEN),
             ("PostgreSQL MCP","DB 效能診斷",C_GOLD),
             ("PR Hook MCP","Code Review 觸發",C_RED)]
for i,(svc,desc,col) in enumerate(left_svcs):
    sy = 1.5 + i*1.55
    rect(s, 0.4, sy, 3.9, 1.3, fill=C_NAVY)
    rect(s, 0.4, sy, 0.1, 1.3, fill=col)
    txt(s, svc,  0.62, sy+0.12, 3.5, 0.45, size=13, bold=True, color=col)
    txt(s, desc, 0.62, sy+0.6,  3.5, 0.38, size=11.5, color=C_LGRAY)
    # connector
    txt(s, "→", 4.32, sy+0.48, 0.4, 0.35, size=14, color=C_MGRAY, align=PP_ALIGN.CENTER)

# Right services
right_svcs = [("Nagios / Prometheus","維運監控整合",C_ACCENT),
              ("AEM 測試 MCP","元件自動覆蓋",RGBColor(0xA0,0x55,0xFF)),
              ("維運推論 MCP","多工具串聯推論",C_GREEN)]
for i,(svc,desc,col) in enumerate(right_svcs):
    sy = 1.5 + i*1.55
    rect(s, 9.05, sy, 3.9, 1.3, fill=C_NAVY)
    rect(s, 9.05, sy, 0.1, 1.3, fill=col)
    txt(s, svc,  9.28, sy+0.12, 3.5, 0.45, size=13, bold=True, color=col)
    txt(s, desc, 9.28, sy+0.6,  3.5, 0.38, size=11.5, color=C_LGRAY)
    txt(s, "←", 8.68, sy+0.48, 0.4, 0.35, size=14, color=C_MGRAY, align=PP_ALIGN.CENTER)

bottom_bar(s, right="5 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 6 ── Cases 1+2
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "從被動應對，到主動預防",
        "兩個核心痛點，AI 都給出了系統性解法，不是一次性修補", "應用案例")

case_card(s, "01", "UI 自動化回歸測試平台",
          "測試腳本零散、人工維護成本高、回歸覆蓋不穩定",
          "Playwright + MCP Server 建立可擴充持續回歸測試平台\n處理 OTP / Captcha / 2FA 後全站探勘自動執行",
          "零散腳本 → 結構化持續平台\n覆蓋率提升且可自動觸發，降低導入成本",
          0.4, 1.2, 6.1, 5.5, C_GREEN)

case_card(s, "02", "PostgreSQL 效能優化",
          "DB 問題靠人工排查，知識不傳承，每次都重新來過",
          "AI 管線：快照截取 → 報表解析 → 智能過濾（容量>1GB 或 Top5 慢查詢）→ AI 診斷建議 → 視覺化監控",
          "被動滅火 → 主動預警\n資深工程師知識封裝傳承，不再依賴個人",
          6.85, 1.2, 6.1, 5.5, C_GOLD)

bottom_bar(s, right="6 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 7 ── Cases 3+4
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "AI 進入審核與維運核心",
        "高風險 PR 直接退回，維運推論秒出結果，AI 已是品質與穩定的守門人", "應用案例")

case_card(s, "03", "PR Code Review「無情審查員」",
          "人工 Code Review 品質不穩定，高風險變更難以即時識別",
          "AI Agent 雙層審核：\n第一道：開發 Reviewer（需求範圍與功能正確性）\n第二道：維運 Reviewer（穩定性、可用性、風險防禦）",
          "輸出 High / Medium / Low 風險報告\nHigh Risk PR 直接退回修正，不進人工審核流程",
          0.4, 1.2, 6.1, 5.5, C_RED)

case_card(s, "04", "DSP 維運助手",
          "維運問題需跨系統查詢，耗時且依賴資深人員經驗",
          "輸入門號 → AI 自動串聯多 MCP 工具推論：\nUser ID 查詢 → 交易紀錄 → Nagios/Prometheus 監控 → 綜合風險分析",
          "秒級輸出完整風險結論\n降低維運人力門檻，知識民主化",
          6.85, 1.2, 6.1, 5.5, C_ACCENT)

bottom_bar(s, right="7 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 8 ── AEM Migration
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "規模龐大到人工無法完整覆蓋的專案，AI 做到了 100%",
        "AEM Cloud Migration：AI 能力的規模壓力測試，通過了", "大型專案驗證")

# 3 big numbers
for i,(num,unit,desc,col) in enumerate([
    ("292","個元件","全部完成升級",C_ACCENT),
    ("27,709","個網頁","全站搬移完成",C_GOLD),
    ("100,000+","件素材","完整遷移覆蓋",C_RED),
]):
    nx = 0.5 + i*4.25
    rect(s, nx, 1.3, 4.0, 1.85, fill=C_NAVY)
    rect(s, nx, 1.3, 4.0, 0.08, fill=col)
    txt(s, num,  nx+0.2, 1.42, 3.7, 1.0, size=46, bold=True, color=col, align=PP_ALIGN.CENTER)
    txt(s, unit, nx+0.2, 2.42, 3.7, 0.38, size=16, bold=True, align=PP_ALIGN.CENTER)
    txt(s, desc, nx+0.2, 2.82, 3.7, 0.3,  size=11.5, color=C_LGRAY, align=PP_ALIGN.CENTER)

# Before / After
rect(s, 0.5, 3.45, 5.9, 3.0, fill=RGBColor(0x18,0x10,0x10))
rect(s, 0.5, 3.45, 5.9, 0.07, fill=C_DGRAY)
txt(s, "導入 AI 前", 0.7, 3.55, 5.5, 0.4, size=14, bold=True, color=C_MGRAY)
for i,pt in enumerate(["測試進度嚴重落後","人工覆蓋率無法達標","細微樣式差異難以偵測"]):
    txt(s, f"✕  {pt}", 0.7, 4.05+i*0.65, 5.5, 0.52, size=13, color=RGBColor(0xEF,0x44,0x44))

rect(s, 6.9, 3.45, 5.9, 3.0, fill=RGBColor(0x0A,0x18,0x12))
rect(s, 6.9, 3.45, 5.9, 0.07, fill=C_GREEN)
txt(s, "導入 AI 後", 7.1, 3.55, 5.5, 0.4, size=14, bold=True, color=C_GREEN)
for i,pt in enumerate(["100% FETnet 全站自動化覆蓋","AI 自動分析異常原因","細微樣式差異自動偵測"]):
    txt(s, f"✓  {pt}", 7.1, 4.05+i*0.65, 5.5, 0.52, size=13, color=C_GREEN)

txt(s, "→", 6.35, 4.5, 0.6, 0.6, size=28, bold=True, color=C_RED, align=PP_ALIGN.CENTER)

bottom_bar(s, right="8 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 9 ── Quantified Results Dashboard
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
top_bar(s, "數字不說謊，成果已驗證",
        "每一項數字背後都是真實交付，不是 POC，不是實驗室數據", "量化成果")

dashboard = [
    ("12 → 105","AI 社群規模","三個月成長 8.75 倍",C_ACCENT),
    ("20+","Skills 封裝","可重用 AI 能力模組已上線",C_GOLD),
    ("10+","MCP Services","平台運行中的工具服務",C_RED),
    ("5 / 5","SDLC 覆蓋","需求到維運，全環節有 Agent",C_GREEN),
    ("100%","自動化覆蓋","AEM Migration 全站達成",RGBColor(0xA0,0x55,0xFF)),
    ("2 層","審核機制","PR Code Review 雙層把關",C_ACCENT),
]
for i,(v,l,d,c) in enumerate(dashboard):
    col=i%3; row=i//3
    sx=0.42+col*4.3; sy=1.32+row*2.52
    rect(s, sx, sy, 4.05, 2.28, fill=C_NAVY)
    rect(s, sx, sy, 4.05, 0.07, fill=c)
    txt(s, v, sx+0.2, sy+0.1,  3.7, 0.95, size=40, bold=True, color=c)
    txt(s, l, sx+0.2, sy+1.05, 3.7, 0.38, size=14, bold=True)
    txt(s, d, sx+0.2, sy+1.48, 3.7, 0.48, size=11, color=C_LGRAY)

bottom_bar(s, right="9 / 10")


# ══════════════════════════════════════════════════════════════
# Slide 10 ── Vision & CTA
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
bg(s)
rect(s, 0, 0, 13.33, 7.5, fill=C_DARK)
rect(s, 0, 0, 0.2, 7.5, fill=C_RED)

txt(s, "基礎已立", 0.45, 0.3,  12, 0.9, size=40, bold=True, color=C_WHITE)
txt(s, "下一步等您點頭", 0.45, 1.18, 12, 0.9, size=40, bold=True, color=C_GOLD)
hline(s, 0.45, 2.18, 5.5)

# Done / Next columns
rect(s, 0.45, 2.38, 5.85, 3.85, fill=C_NAVY)
rect(s, 0.45, 2.38, 5.85, 0.07, fill=C_GREEN)
txt(s, "✅  已完成（基礎紮實）", 0.65, 2.48, 5.5, 0.45, size=14, bold=True, color=C_GREEN)
for i,pt in enumerate(["AI 架構驗證（LLM+Skills+MCP）",
                        "MCPHub 平台建置與上線",
                        "四項可量化 AI 應用案例交付",
                        "SDLC 五環節 AI Agent 覆蓋",
                        "社群建立（12 → 105 人）"]):
    txt(s, f"  ·  {pt}", 0.65, 3.05+i*0.6, 5.5, 0.52, size=12.5, color=C_WHITE)

rect(s, 6.65, 2.38, 6.25, 3.85, fill=C_NAVY)
rect(s, 6.65, 2.38, 6.25, 0.07, fill=C_RED)
txt(s, "🎯  下一步（請高層支持）", 6.85, 2.48, 5.9, 0.45, size=14, bold=True, color=C_RED)
nexts = [
    ("MCPHub 跨部門開放","將平台能力擴散至其他業務單位"),
    ("AI 技術學院正式建制","從自發社群升格為部門級常設機構"),
    ("SDLC AI 化納入集團規範","讓 AI 工程標準成為全集團基線"),
]
for i,(title,desc) in enumerate(nexts):
    ty = 3.05+i*1.1
    tag(s, title, 6.85, ty, bg_c=C_RED, size=11)
    txt(s, desc, 6.85, ty+0.38, 5.9, 0.55, size=12, color=C_LGRAY)

# CTA
rect(s, 0.45, 6.38, 12.45, 0.5, fill=C_RED)
txt(s, "時間窗口明確，競爭優勢正在形成，現在是擴大投資的最佳時機",
    0.7, 6.42, 12.0, 0.4, size=15, bold=True, align=PP_ALIGN.CENTER)

bottom_bar(s, right="10 / 10")


# ── Save ──────────────────────────────────────────────────────
out = "/home/user/Moji/DTIS_AI_Academy_Report.pptx"
prs.save(out)
print(f"✅  Saved: {out}")
