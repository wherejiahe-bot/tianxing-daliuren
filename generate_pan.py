#!/usr/bin/env python
"""天星2.0 排盘HTML - 精确呈现四课盘面 + 天盘5干"""
import openpyxl
SRC = 'F:/易学赚钱/大六壬/天星2.0.xlsx'
wb = openpyxl.load_workbook(SRC, data_only=True)
s1 = wb['天星1.0']   # 有缓存值的输入区
s2 = wb['天星2.0']   # 静态排盘结果

def rd(sheet, r, col):
    col_idx = 0
    for ch in col: col_idx = col_idx * 26 + (ord(ch) - 64)
    v = sheet.cell(row=r, column=col_idx).value
    if v is None: return ''
    s = str(v).strip()
    return '' if s in ('None', '#N/A', '') else s

def v(col, r):
    """读值，优先天星2.0，回退天星1.0"""
    return rd(s2, r, col) or rd(s1, r, col)

H = []
ap = H.append

ap("""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>天星2.0 · 大六壬排盘</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#f5efe0;color:#2c2c2c;font-family:'Microsoft YaHei','PingFang SC','Noto Sans SC',sans-serif;display:flex;justify-content:center;padding:0}
.wrap{max-width:900px;margin:0;padding:8px 10px;width:100%;background:#faf5ee}

/* ── 顶部 ── */
.top{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.top h1{font-size:16px;color:#8b1a1a;font-weight:bold;letter-spacing:2px}
.top .sub{font-size:11px;color:#888}
.top .hint{font-size:10px;color:#aaa;text-align:right;line-height:1.4}

/* ── 双列四柱 ── */
.sz-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:6px}
.sz-box{background:#fff;border:1px solid #d0c0b0;border-radius:6px;padding:8px}
.sz-box .st{font-size:12px;color:#8b1a1a;font-weight:bold;border-bottom:1px solid #e0d0c0;padding-bottom:3px;margin-bottom:5px}
.sz-row{display:flex;align-items:center;gap:5px;margin-bottom:2px;font-size:13px}
.sz-row .l{color:#999;min-width:32px;text-align:right;font-size:11px}
.sz-row .sv{color:#1a1a2e;font-weight:bold}
.sz-row .sv.r{color:#8b1a1a}
.sz-row input{background:#fafafa;border:1px solid #ddd;border-radius:3px;padding:2px 5px;font-size:13px;width:50px;font-family:inherit}
.sz-note{font-size:10px;color:#bbb;margin-top:3px;line-height:1.3}

/* ── 排盘结果 ── */
.rp{background:#fcf0e0;border:1px solid #c0392b;border-radius:6px;padding:5px 12px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center}
.rp .r{font-size:10px;color:#8b1a1a}
.rp .v{font-size:18px;font-weight:bold;color:#1a1a2e;letter-spacing:3px}
.rp .s{font-size:11px;color:#666}

/* ── 神煞标签 ── */
.tag-b{display:flex;flex-wrap:wrap;gap:3px;margin-bottom:5px}
.tag{background:#fff;border:1px solid #e0d0c0;border-radius:3px;padding:1px 5px;font-size:10px;white-space:nowrap}
.tag .tl{color:#999}
.tag .tv{color:#333;font-weight:bold}

/* ── 中黄人遁 ── */
.zl{display:flex;gap:5px;margin-bottom:6px}
.zl-item{flex:1;background:#fff;border:1px solid #e0d0c0;border-radius:4px;padding:3px 6px;text-align:center;font-size:12px}
.zl-item .zll{color:#999;font-size:10px}
.zl-item .zlv{color:#8b1a1a;font-weight:bold;font-size:13px}

/* ── 四课核心盘面 ── */
.ke-table{border:2px solid #8b1a1a;border-radius:6px;overflow:hidden;margin-bottom:6px;background:#fff}
.ke-tr{display:grid;grid-template-columns:repeat(4,1fr)}
.ke-tr.header{border-bottom:1px solid #d0c0b0}
.ke-tr.header .kc{border-right:1px solid #e0d0c0;padding:4px;text-align:center;font-size:11px;color:#8b1a1a;font-weight:bold}
.ke-tr.header .kc:last-child{border-right:none}
.ke-tr.data .kc{padding:3px 4px;text-align:center;border-right:1px solid #f0e0d0;border-bottom:1px solid #f0e0d0}
.ke-tr.data .kc:last-child{border-right:none}
.ke-tr.data:last-child .kc{border-bottom:none}

.kc .tp{font-size:17px;font-weight:bold;color:#1a1a2e;letter-spacing:1px}
.kc .bt{font-size:13px;color:#555;margin-top:1px}
.kc .god{font-size:12px;margin-top:1px;font-weight:bold}
.kc .det{font-size:9px;color:#888;margin-top:1px;line-height:1.2}
.kc .sx{font-size:9px;color:#aaa;margin-top:1px}

/* 天将色 + 地支 */
.g-she{color:#6a1b9a}.g-zhu{color:#c62828}.g-he{color:#2e7d32}
.g-gou{color:#e65100}.g-qing{color:#1565c0}.g-kong{color:#546e7a}
.g-bai{color:#37474f}.g-chang{color:#00695c}.g-xuan{color:#1a237e}
.g-yin{color:#4a148c}.g-hou{color:#880e4f}.g-gui{color:#f9a825}

/* ── 天盘五干 ── */
.wg{background:#fff0e0;border:1px solid #c0392b;border-radius:4px;padding:5px 8px;margin-bottom:5px;display:flex;flex-wrap:wrap;gap:4px;align-items:center}
.wg .l{font-size:10px;color:#8b1a1a;font-weight:bold}
.wg .item{background:#fff;border:1px solid #e0c0a0;border-radius:3px;padding:1px 6px;font-size:11px}
.wg .item .il{color:#999}
.wg .item .iv{color:#8b1a1a;font-weight:bold}

/* ── 五行 ── */
.wx{display:flex;gap:3px;margin:4px 0}
.wx span{font-size:10px;padding:1px 5px;border-radius:3px}
.wx-huo{background:#fbe9e7;color:#c62828}.wx-tu{background:#fff3e0;color:#e65100}
.wx-mu{background:#e8f5e9;color:#2e7d32}.wx-shui{background:#e3f2fd;color:#1565c0}
.wx-jin{background:#f3e5f5;color:#6a1b9a}

/* ── 神煞明细 ── */
.ss{border-top:1px solid #e0d0c0;padding-top:5px;margin-top:4px}
.ss-t{font-size:11px;color:#8b1a1a;font-weight:bold;margin-bottom:3px}
@media(max-width:640px){.sz-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrap">
""")

# ─── 标题 ───
ap('<div class="top">')
ap('<div><h1>天星2.0 · 大六壬</h1><div class="sub">九天玄数 · 太极辨证法 · 神煞集成</div></div>')
ap('<div class="hint">三传: 中黄 旬遁 月将<br>贵神 人遁 天遁</div></div>')

# ─── 自选/排盘四柱 ───
ap('<div class="sz-grid">')
# 左：自选
ap('<div class="sz-box"><div class="st">自选 四柱</div>')
ap(f'<div class="sz-row"><span class="l">年</span><input value="{v("P",2)}{v("P",3)}" placeholder="例:2024">'
   f'<span class="l">月将</span><span class="sv">{v("I",1)}</span></div>')
ap(f'<div class="sz-row"><span class="l">性别</span><span class="sv">{v("J",2)}</span>'
   f'<span class="l">年份</span><span class="sv">{v("K",2)}</span></div>')
ap(f'<div class="sz-row"><span class="l">本命</span><span class="sv">{v("L",2)}</span>'
   f'<span class="l">行年</span><span class="sv">{v("M",2) or "—"}</span></div>')
ap('<div class="sz-note">删掉自选即为当下正时课</div></div>')

# 右：排盘
ap('<div class="sz-box"><div class="st">排盘 四柱</div>')
ap(f'<div class="sz-row"><span class="l">年</span><span class="sv r">{v("P",2)}{v("P",3)}</span>'
   f'<span class="l">月</span><span class="sv r">{v("Q",2)}{v("Q",3)}</span></div>')
ap(f'<div class="sz-row"><span class="l">日</span><span class="sv r">{v("R",2)}{v("R",3)}</span>'
   f'<span class="l">时</span><span class="sv r">{v("S",2)}{v("S",3)}</span></div>')
ap(f'<div class="sz-row"><span class="l">性别</span><span class="sv">{v("J",3)}</span>'
   f'<span class="l">本命</span><span class="sv">{v("L",3)}</span></div>')
ap('<div class="sz-note">当下正时 : 丙午 癸巳 壬寅 丁未</div></div></div>')

# ─── 排盘结果 ───
ap(f'<div class="rp"><div><div class="r">天星排盘 · 四柱</div>'
   f'<div class="s">{v("P",2)}{v("P",3)}年 {v("Q",2)}{v("Q",3)}月 {v("R",2)}{v("R",3)}日 {v("S",2)}{v("S",3)}时</div></div>'
   f'<div class="v">{v("P",2)}{v("Q",2)}{v("R",2)}{v("S",2)}　{v("P",3)}{v("Q",3)}{v("R",3)}{v("S",3)}</div></div>')

# ─── 神煞标签 ───
ap('<div class="tag-b">')
for label, col, r in [('日德','U',2),('大耗','W',2),('将星','Y',2),('月德','U',3),('天赦','W',3),('文昌','Y',3),
                        ('天德','U',5),('孤辰','W',5),('劫煞','Y',5),('天喜','W',3),('生气','Y',3)]:
    val = v(col, r)
    if val and val != label:
        ap(f'<span class="tag"><span class="tl">{label}</span><span class="tv"> {val}</span></span>')
ap('</div>')

# ─── 中黄天遁 ───
ap('<div class="zl">')
ap(f'<div class="zl-item"><div class="zll">中黄</div><div class="zlv">{v("O",8)[:12] or "—"}</div></div>')
ap(f'<div class="zl-item"><div class="zll">人遁</div><div class="zlv">{v("I",8) or "—"}</div></div>')
ap(f'<div class="zl-item"><div class="zll">天遁</div><div class="zlv">{v("K",8) or "—"}</div></div>')
ap(f'<div class="zl-item"><div class="zll">旬遁</div><div class="zlv">{v("AI",5)}{v("AJ",5) or "—"}</div></div>')
ap('</div>')

# ─── 四课核心盘面 ───
ap('<div class="ke-table">')
# 表头
ap('<div class="ke-tr header">')
for label in ['第一课','第二课','第三课','第四课']:
    ap(f'<div class="kc">{label}</div>')
ap('</div>')

# 数据行
ke_keys = ['D','H','L','P']

# 行1: 天盘天干
ap('<div class="ke-tr data">')
for start in ke_keys:
    tg = v(start, 7)
    tz = v(chr(ord(start)+2), 7)
    ap(f'<div class="kc"><div class="tp">{tg} {tz}</div>')
    ap(f'<div class="det">天盘天干·地支</div></div>')
ap('</div>')

# 行2: 地盘天干
ap('<div class="ke-tr data">')
for start in ke_keys:
    dg = v(start, 4)
    # 取row 5的地支
    dz_row5 = v(chr(ord(start)+2), 5)
    ap(f'<div class="kc"><div class="bt">{dg} {dz_row5}</div>')
    ap(f'<div class="det">地盘天干·地支</div></div>')
ap('</div>')

# 行3: 神将
ap('<div class="ke-tr data">')
for start in ke_keys:
    god = v(chr(ord(start)+1), 7)
    dg_god = v(chr(ord(start)+1), 4)
    
    cls = ''
    for kw, c in [('螣','she'),('朱','zhu'),('合','he'),('勾','gou'),('青','qing'),
                   ('空','kong'),('白','bai'),('常','chang'),('玄','xuan'),('阴','yin'),('后','hou'),('贵','gui')]:
        if kw in god: cls = f'g-{c}'
    
    ap(f'<div class="kc"><div class="god {cls}">{god}</div>')
    ap(f'<div class="det">天将: {dg_god}</div></div>')
ap('</div>')

# 行4: 长生状态 + 细节
ap('<div class="ke-tr data">')
for start in ke_keys:
    detail = v(chr(ord(start)+3), 7).replace('  ',' ').strip()
    shenx = v(chr(ord(start)+2), 6)  # 长生状态在Row6
    ap(f'<div class="kc"><div class="sx">{detail[:20]}</div>')
    if shenx: ap(f'<div class="det">长生: {shenx}</div>')
    ap('</div>')
ap('</div>')

# 行5: 十二神将名
ap('<div class="ke-tr data">')
for start in ke_keys:
    god_name = v(chr(ord(start)+3), 5)  # 神将名(胜光, 小吉, 传送, 从魁)
    ap(f'<div class="kc"><div class="det" style="color:#666;font-size:11px">{god_name}</div></div>')
ap('</div>')

ap('</div>')  # end ke-table

# ─── 天盘五干运算 ───
ri_gan = v('D',4)  # 日干 = 丙

# 四课天盘天干
tp_tg = [v(c, 7) for c in ke_keys]
# 四课地盘天干
dp_tg = [v(c, 4) for c in ke_keys]

# 五鼠遁: 丙辛从戊起
# 日干=丙, 五鼠遁起点: 戊子
# 根据截图展示5个核心天干
wushu_info = "丙日 · 五鼠遁: 戊子起"

ap('<div class="wg">')
ap('<span class="l">天盘五干</span>')

# 日干作为核心
ap(f'<span class="item"><span class="il">日干</span> <span class="iv">{ri_gan}</span></span>')

# 四课天盘天干
for i, tg in enumerate(tp_tg):
    ap(f'<span class="item"><span class="il">课{i+1}天盘</span> <span class="iv">{tg}</span></span>')

# 四课地盘天干
for i, dg in enumerate(dp_tg):
    ap(f'<span class="item"><span class="il">课{i+1}地盘</span> <span class="iv">{dg}</span></span>')

ap(f'<span class="item" style="background:#fcf0e0"><span class="il">五鼠遁</span> <span class="iv">{wushu_info}</span></span>')
ap('</div>')

# ─── 五行状态 ───
ap('<div class="wx">')
ap(f'<span class="wx-huo">日干: {ri_gan}</span>')
ap('<span class="wx-huo">旺: 火旺</span><span class="wx-tu">相: 土相</span>')
ap('<span class="wx-mu">休: 木休</span><span class="wx-shui">囚: 水囚</span>')
ap('<span class="wx-jin">死: 金死</span>')
ap('</div>')

# ─── 神煞明细 ───
ap('<div class="ss"><div class="ss-t">神煞明细</div><div class="tag-b">')
for r in [1,2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19]:
    for col in ['V','W','X','Y','Z','AA','AB','AC','AD','AE','AF']:
        val = v(col, r)
        label = v(col, 1)
        if val and label and val != label:
            ap(f'<span class="tag"><span class="tl">{label}</span> <span class="tv">{val}</span></span>')
ap('</div></div>')

ap('<div style="text-align:center;color:#ccc;font-size:10px;margin-top:8px;padding:6px 0">天星2.0 · 霎哈嘉瑜伽修习者</div>')
ap('</div></body></html>')

out = 'G:/WorkBuddy/码农龙虾/tianxing-daliuren/index.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(H))
print(f'Done: {out} ({sum(len(h) for h in H)}b)')
