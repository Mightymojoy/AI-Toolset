# -*- coding: utf-8 -*-
"""
AI 工具集门户生成器（未来玻璃版 / glass build）
读取同目录下「AI 工具集.xlsx」的 Sheet1，自动生成玻璃拟态风格 index.html。

表格结构（Sheet1）：
  - 分组行：A列有值、B/C为空  -> 来源分组（如「自建网页」「飞书文档」），仅作诊断信息
  - 表头行：含「AI智能体名称 / 对应链接 / 分类」等标题词 -> 跳过
  - 数据行：三列都有值 -> {name, url, cat}
  - 锁定的工具：名称含「暂未解锁 / 待解锁」或链接为空 -> 标为 locked（灰显、不跳外链）

设计说明：
  - 视觉/中英双语/I18N/动画全部来自 portal_template.html（玻璃版模板），本脚本只负责填数据
  - 每个工具的中英描述与图标按「工具名」映射到 TOOL_META；分类的中英名/图标/主题色来自 CAT_ORDER
  - xlsx 仅作为工具清单来源，结构不变；新增工具直接在表格里加一行，重跑脚本即可刷新门户
"""
import openpyxl
import json
import os
import re
import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(SCRIPT_DIR, "AI 工具集.xlsx")
TEMPLATE = os.path.join(SCRIPT_DIR, "portal_template.html")
OUT = os.path.join(SCRIPT_DIR, "index.html")


def _norm(s):
    """归一化：去空格 + 去全角/半角括号及其中内容。"""
    s = s.replace(" ", "")
    s = re.sub(r"[（(].*?[)）]", "", s)
    return s.strip()


def _clean_name(name):
    """清洗工具名：去掉「（暂未解锁）」等标记。"""
    n = re.sub(r"[（(]?(暂未解锁|待解锁)[)）]?", "", name)
    return n.strip()


def _is_locked(name, url):
    if "暂未解锁" in name or "待解锁" in name:
        return True
    if not url or url.strip().lower() in ("#", "locked", "todo"):
        return True
    return False


# ---------- 分类定义（显示顺序 / 中英名 / 图标 / 主题色）----------
CAT_ORDER = [
    ("数据看板",     "dash",  "Dashboards",        "dashboard", "#7C3AED"),
    ("知识库与文档", "docs",  "Docs & Knowledge",  "bookOpen",  "#EC4899"),
    ("内容创作工具", "create", "Content Tools",     "pen",      "#06B6D4"),
    ("数据表",       "sheet", "Data Tables",        "table",    "#7C3AED"),
    ("AI资讯",       "news",  "AI News",            "news",     "#EC4899"),
]
CAT_KEYS = [c[1] for c in CAT_ORDER]
# 分类名归一化（去空格、去全角括号）-> key
CAT_ALIAS = {}
for raw, key, *_ in CAT_ORDER:
    CAT_ALIAS[_norm(raw)] = key
CAT_ALIAS["ai资讯"] = "news"   # xlsx 无空格写法兜底

# 分类兜底描述（当某工具名不在 TOOL_META 时使用）
CAT_DESC = {
    "dash":   {"zh": "按系列 / 渠道追踪各项数据", "en": "Track data by series and channel"},
    "docs":   {"zh": "产品资料与知识沉淀",       "en": "Product docs and knowledge"},
    "create": {"zh": "内容创作提效工具",         "en": "Content creation tools"},
    "sheet":  {"zh": "飞书数据明细表",           "en": "Feishu data sheets"},
    "news":   {"zh": "AI 行业每日速览",          "en": "Daily AI industry digest"},
}
CAT_ICON = {c[1]: c[3] for c in CAT_ORDER}

# ---------- 工具富文本映射（key = 清洗后的工具名）----------
# 值：(zh 描述, en 描述, 图标)
TOOL_META = {
    "ITO产品销售分析看板":      ("按系列 SKU 拆解销售表现",            "Sales performance by series and SKU",            "barChart"),
    "ITO电商渠道业绩看板":      ("9 渠道月度目标与实际追踪",          "9 channels: monthly target vs actual",          "trending"),
    "抖音主播数据看板":         ("三店主播实时直播数据",               "Live-stream data across 3 stores",               "monitor"),
    "抖音主播数据趋势分析":     ("主播历史趋势与对比",                 "Host trends and comparison",                     "activity"),
    "抖音店铺数据看板":         ("抖音自营店铺经营概览",               "Douyin store operations overview",               "store"),
    "京东毛保监控":             ("京东自营毛利保护每日看板",           "JD self-operated margin protection daily",       "shield"),
    "店铺优惠券数据监控":       ("优惠券投放与核销监控",               "Coupon delivery and redemption monitor",         "ticket"),
    "ITO库存看板":              ("全渠道库存水位与周转",               "Omni-channel inventory and turnover",            "package"),
    "ITO产品介绍":              ("全系产品资料与卖点",                 "Full product catalog and selling points",        "book"),
    "ITO新品介绍":              ("新品上市资料中心",                   "New product launch resource center",            "sparkles"),
    "新品上线进度概览":         ("新品上市流程进度看板",               "New product launch progress board",              "kanban"),
    "ITO AI AGENT 测试":        ("AI 评分与 Agent 测试台",             "AI scoring and agent testing bench",             "bot"),
    "AI培训视频讲解":           ("AI 实操培训视频（待解锁）",          "AI hands-on training videos (locked)",           "play"),
    "ITO内容短视频脚本生成器":   ("抖音短视频脚本一键生成",             "One-click Douyin short-video scripts",           "video"),
    "ITO摄影提示词精准控制生成": ("产品摄影提示词工坊",                 "Product photography prompt studio",              "camera"),
    "每日库存表":               ("飞书每日库存明细表",                 "Feishu daily inventory sheet",                    "table"),
    "京东毛保":                 ("京东毛保明细数据表",                  "JD margin-protection data sheet",                 "table"),
    "电商渠道业绩":             ("渠道业绩明细表",                     "Channel performance sheet",                       "table"),
    "电商渠道商品":             ("渠道在售商品表",                     "Channel product listing sheet",                   "table"),
    "新品上市全流程":           ("新品上市流程管理表",                 "New product launch workflow sheet",               "table"),
    "AI HOT 每日晨报":          ("每日 AI 行业早报",                   "Daily AI industry morning brief",                 "zap"),
}


def main():
    if not os.path.exists(SRC):
        raise SystemExit("未找到数据源：%s" % SRC)
    if not os.path.exists(TEMPLATE):
        raise SystemExit("未找到模板：%s" % TEMPLATE)

    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Sheet1"]

    HEADER_WORDS = {"AI智能体名称", "名称", "工具名", "工具名称", "对应链接", "链接", "分类"}
    raw_entries = []          # 顺序保留的原始数据行
    skipped_header = 0
    skipped_group = 0
    skipped_incomplete = []
    group = None

    for row in ws.iter_rows(values_only=True):
        cells = list(row) + [None, None, None]
        a, b, c = cells[0], cells[1], cells[2]
        if a is None and b is None and c is None:
            continue
        a_s = str(a).strip() if a is not None else ""
        b_s = str(b).strip() if b is not None else ""
        c_s = str(c).strip() if c is not None else ""
        if not a_s:
            continue
        if a_s in HEADER_WORDS:
            skipped_header += 1
            continue
        if a_s and not b_s and not c_s:   # 分组行
            group = a_s
            skipped_group += 1
            continue
        if a_s and b_s and c_s:
            raw_entries.append({"group": group or "", "name": a_s, "url": b_s, "cat": c_s})
        else:
            skipped_incomplete.append(a_s)

    if not raw_entries:
        raise SystemExit("未从表格中解析到任何工具条目，请检查 Sheet1 结构。")

    # ---------- 构建 TOOLS ----------
    tools = []
    unknown_cats = []
    for e in raw_entries:
        clean = _clean_name(e["name"])
        locked = _is_locked(e["name"], e["url"])
        cat_key = CAT_ALIAS.get(_norm(e["cat"]))
        if cat_key is None:
            unknown_cats.append(e["cat"])
            cat_key = CAT_KEYS[0]   # 兜底，避免整页崩
        meta = TOOL_META.get(clean)
        if meta:
            desc_zh, desc_en, icon = meta
        else:
            d = CAT_DESC.get(cat_key, CAT_DESC[CAT_KEYS[0]])
            desc_zh, desc_en = d["zh"], d["en"]
            icon = CAT_ICON.get(cat_key, "box")
        tools.append({
            "cat": cat_key,
            "name": clean,
            "url": ("" if locked else e["url"]),
            "icon": icon,
            "desc": {"zh": desc_zh, "en": desc_en},
            "locked": locked,
        })

    # ---------- 构建 CATS（仅包含实际出现的分类，保持显示顺序）----------
    seen = []
    for k in CAT_KEYS:
        if any(t["cat"] == k for t in tools):
            seen.append(k)
    cats = []
    for raw, key, en_label, icon, accent in CAT_ORDER:
        if key in seen:
            cats.append({
                "key": key,
                "label": {"zh": raw, "en": en_label},
                "icon": icon,
                "accent": accent,
            })

    cats_json = json.dumps(cats, ensure_ascii=False)
    tools_json = json.dumps(tools, ensure_ascii=False)
    gen_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    with open(TEMPLATE, "r", encoding="utf-8") as f:
        html = f.read()
    html = (html
            .replace("window.__CATS__", cats_json)
            .replace("window.__TOOLS__", tools_json)
            .replace("__GEN__", gen_time))

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    print("OK 解析到 %d 个工具，分类 %d 个" % (len(tools), len(cats)))
    print("  - 表头跳过 %d 行，分组标题 %d 个" % (skipped_header, skipped_group))
    if skipped_incomplete:
        print("  - 警告：%d 行只有名称、缺少链接或分类，未收录 -> %s" % (len(skipped_incomplete), skipped_incomplete))
    if unknown_cats:
        print("  - 警告：%d 个未知分类已兜底到首类 -> %s" % (len(unknown_cats), sorted(set(unknown_cats))))
    print("  - 锁定工具：%s" % ([t["name"] for t in tools if t["locked"]] or "无"))
    print("已生成 -> %s" % OUT)


if __name__ == "__main__":
    main()
