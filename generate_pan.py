#!/usr/bin/env python
"""天星2.0 排盘HTML生成器 - 从两个sheet读取数据"""
import openpyxl
SRC = 'F:/易学赚钱/大六壬/天星2.0.xlsx'

wb = openpyxl.load_workbook(SRC, data_only=True)
s1 = wb['天星1.0']  # 有缓存的公式值
s2 = wb['天星2.0']  # 静态复制（有排盘结果）

def c(sheet, r, col_str):
    col_idx = 0
    for ch in col_str:
        col_idx = col_idx * 26 + (ord(ch) - 64)
    v = sheet.cell(row=r, column=col_idx).value
    if v is None: return ''
    s = str(v).strip()
    return '' if s in ('None', '#N/A', '', ' ') else s

def v(col_str, r, prefer_s2=True):
    """读取值，优先天星2.0（静态），回退天星1.0（缓存）"""
    v2 = c(s2, r, col_str)
    return v2 if v2 else c(s1, r, col_str)

H = []
H.append("""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>天星2.0 · 大六壬排盘</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Microsoft YaHei','PingFang SC','Noto Sans SC',sans-serif;
  background:#f0ebe3;color:#2c2c2c;padding:0;display:flex;justify-content:center}
.container{max-width:860px;margin:0 auto;padding:12px;width:100%}

/* ── 顶部：双列四柱 ── */
.header-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
.header-row h1{font-size:18px;color:#a02020;font-weight:bold;letter-spacing:2px}
.header-row .sub{font-size:12px;color:#666;margin-top:1px}

.sizhu-panel{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px}
.sizhu-box{background:#fff8f0;border:1px solid #dcc;border-radius:6px;padding:8px}
.sizhu-box .stitle{font-size:12px;color:#a02020;font-weight:bold;border-bottom:1px solid #e0d0c0;padding-bottom:4px;margin-bottom:6px}
.sizhu-row{display:flex;align-items:center;gap:6px;margin-bottom:3px;font-size:13px}
.sizhu-row .sl{color:#888;min-width:36px;text-align:right;font-size:11px}
.sizhu-row .sv{color:#1a1a2e;font-weight:bold;min-width:40px}
.sizhu-row .sv.gold{color:#a02020}
.sizhu-row input{background:#fff;border:1px solid #ddd;border-radius:3px;padding:3px 6px;font-size:13px;width:60px;font-family:inherit}
.sizhu-row input:focus{outline:none;border-color:#a02020}
.sizhu-hint{font-size:10px;color:#aaa;line-height:1.4;margin-top:4px;padding-left:42px}

/* ── 排盘结果 ── */
.result-bar{background:#fff0e0;border:1px solid #c0392b;border-radius:6px;padding:6px 12px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}
.result-bar .rlabel{font-size:11px;color:#a02020}
.result-bar .rval{font-size:20px;font-weight:bold;color:#1a1a2e;letter-spacing:4px}
.result-bar .rsub{font-size:12px;color:#555}

/* ── 四课网格 ── */
.ke-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;margin-bottom:8px}
.ke-cell{background:#fff6e8;border:2px solid #c0392b;border-radius:6px;padding:6px 4px;text-align:center;position:relative}
.ke-label{font-size:9px;color:#a02020;position:absolute;top:-1px;left:5px;font-weight:bold}
.ke-tian{font-size:18px;font-weight:bold;color:#1a1a2e;margin-top:3px;letter-spacing:2px}
.ke-di{font-size:15px;color:#333;margin-top:1px;letter-spacing:2px}
.ke-god{font-size:12px;margin-top:2px;font-weight:bold}
.ke-detail{font-size:10px;color:#555;margin-top:2px}
.god-she{color:#6a1b9a}.god-zhu{color:#c62828}.god-he{color:#2e7d32}
.god-gou{color:#e65100}.god-qing{color:#1565c0}.god-kong{color:#546e7a}
.god-bai{color:#37474f}.god-chang{color:#00695c}.god-xuan{color:#1a237e}
.god-yin{color:#4a148c}.god-hou{color:#880e4f}.god-gui{color:#f9a825}

/* ── 中黄人遁 ── */
.info-row{display:flex;gap:6px;margin-bottom:8px}
.info-cell{flex:1;background:#fff8f0;border:1px solid #ddd;border-radius:4px;padding:4px 8px;text-align:center;font-size:12px}
.info-cell .il{color:#888;font-size:10px}
.info-cell .iv{color:#a02020;font-weight:bold;font-size:14px}

/* ── 神煞标签 ── */
.tag-row{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:6px}
.tag{background:#fff;border:1px solid #ddd;border-radius:3px;padding:1px 6px;font-size:10px}
.tag .tl{color:#999}.tag .tv{color:#333}

/* ── 五行 ── */
.wx{display:flex;gap:4px;margin:4px 0}
.wx span{font-size:10px;padding:1px 6px;border-radius:3px}
.wx-mu{background:#e8f5e9;color:#2e7d32}.wx-huo{background:#fbe9e7;color:#c62828}
.wx-tu{background:#fff3e0;color:#e65100}.wx-jin{background:#f3e5f5;color:#6a1b9a}
.wx-shui{background:#e3f2fd;color:#1565c0}

/* ── 神煞详细区 ── */
.shensha-section{border-top:1px solid #e0d0c0;padding-top:6px;margin-top:4px}
.shensha-title{font-size:11px;color:#a02020;font-weight:bold;margin-bottom:4px}
@media(max-width:640px){.sizhu-panel{grid-template-columns:1fr}.ke-grid{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="container">

<div class="header-row">
  <div><h1>天星2.0</h1><div class="sub">大六壬 · 九天玄数 · 神煞集成</div></div>
  <div style="text-align:right;font-size:11px;color:#888;line-height:1.5">
    三传: 中黄 旬遁 月将<br>贵神 人遁 天遁
  </div>
</div>

<!-- 自选四柱 / 排盘四柱 -->
<div class="sizhu-panel">
<div class="sizhu-box">
<div class="stitle">自选 四柱</div>
""")

# 自选四柱区
H.append(f'<div class="sizhu-row"><span class="sl">年</span><input value="{c(s1,2,"J")}" placeholder="年"><span class="sv gold">{c(s1,2,"K")}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">年</span><span class="sv">{v("P",2)}{v("P",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">月</span><span class="sv">{v("Q",2)}{v("Q",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">日</span><span class="sv">{v("R",2)}{v("R",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">时</span><span class="sv">{v("S",2)}{v("S",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">性别</span><span class="sv">{c(s1,2,"J")}</span><span class="sl">年份</span><span class="sv">{c(s1,2,"K")}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">本命</span><span class="sv">{c(s1,2,"L")}</span><span class="sl">行年</span><span class="sv">{c(s1,2,"M") or "—"}</span></div>')
H.append('<div class="sizhu-hint">自选四柱删掉即为当下正时课</div>')
H.append('</div>')

# 排盘四柱区
H.append('<div class="sizhu-box">')
H.append('<div class="stitle">排盘 四柱</div>')
p_gan = v('P',2)+v('Q',2)+v('R',2)+v('S',2)
p_zhi = v('P',3)+v('Q',3)+v('R',3)+v('S',3)
H.append(f'<div class="sizhu-row"><span class="sl">年</span><span class="sv gold">{v("P",2)}{v("P",3)}</span><span class="sl">月</span><span class="sv gold">{v("Q",2)}{v("Q",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">日</span><span class="sv gold">{v("R",2)}{v("R",3)}</span><span class="sl">时</span><span class="sv gold">{v("S",2)}{v("S",3)}</span></div>')
H.append(f'<div class="sizhu-row"><span class="sl">月将</span><span class="sv">{c(s1,1,"I") or "—"}</span></div>')
# 月将的值可能在别的列
yuejiang = v('I',2) or v('I',3) or c(s1,2,'I')
H.append(f'<div class="sizhu-row"><span class="sl">本命</span><span class="sv">{c(s1,3,"L") or "—"}</span></div>')
H.append('</div></div>')

# 排盘结果展示
H.append(f'<div class="result-bar">')
H.append(f'<div><div class="rlabel">天星排盘</div><div class="rsub">{p_gan[0]}{p_zhi[0]}年 {p_gan[1]}{p_zhi[1]}月 {p_gan[2]}{p_zhi[2]}日 {p_gan[3]}{p_zhi[3]}时</div></div>')
H.append(f'<div class="rval">{p_gan} {p_zhi}</div>')
H.append('</div>')

# 神煞标签行
H.append('<div class="tag-row">')
pairs = [
    ('日德',v('U',2)), ('大耗',v('W',2)), ('将星',v('Y',2)),
    ('月德',v('U',3)), ('天赦',v('W',3)), ('文昌',v('Y',3)),
    ('天德',v('U',5)), ('孤辰',v('W',5)),
    ('天喜',v('W',3)), ('生气',v('Y',3)),
]
for label, val in pairs:
    if val and val != label:
        H.append(f'<span class="tag"><span class="tl">{label}</span> <span class="tv">{val}</span></span>')
H.append('</div>')

# 中黄人遁天遁
H.append('<div class="info-row">')
H.append(f'<div class="info-cell"><div class="il">中黄</div><div class="iv">{v("O",8)[:12] or "—"}</div></div>')
H.append(f'<div class="info-cell"><div class="il">人遁</div><div class="iv">{v("I",8) or "—"}</div></div>')
H.append(f'<div class="info-cell"><div class="il">天遁</div><div class="iv">{v("K",8) or "—"}</div></div>')
H.append(f'<div class="info-cell"><div class="il">旬遁</div><div class="iv">{v("AI",5)}{v("AJ",5) or "—"}</div></div>')
H.append('</div>')

# 四课
H.append('<div class="ke-grid">')
ke_labels = ['第一课','第二课','第三课','第四课']
ke_keys = ['D','H','L','P']

for idx, start in enumerate(ke_keys):
    # Row7 = 天盘upper, Row5 = 神将lower
    tian_gan = v(start, 7)    # 天干上（天盘）
    # 取Row4的第1列作为天干，Row4的第3列作为地支
    di_gan = v(start, 4)      # 天干下（地盘）
    tian_zhi = v(chr(ord(start)+2), 7)  # 地支
    di_god = v(chr(ord(start)+1), 7)    # 神将
    # 地支从Row5取
    di_zhi_val = v(chr(ord(start)+2), 5)  # 地支（从第五行）
    detail = v(chr(ord(start)+3), 7)      # 细节
    detail = detail.replace('  ',' ').strip() if detail else ''
    
    god_class = ''
    for kw, cls in [('螣','she'),('朱','zhu'),('合','he'),('勾','gou'),('青','qing'),
                     ('空','kong'),('白','bai'),('常','chang'),('玄','xuan'),('阴','yin'),('后','hou'),('贵','gui')]:
        if kw in di_god:
            god_class = f'god-{cls}'
            break
    
    H.append('<div class="ke-cell">')
    H.append(f'<div class="ke-label">{ke_labels[idx]}</div>')
    H.append(f'<div class="ke-tian">{tian_gan}　{di_zhi_val}</div>')
    H.append(f'<div class="ke-di">{di_gan}　{di_god}</div>')
    H.append(f'<div class="ke-god {god_class}">{v(chr(ord(start)+1), 4)}</div>')
    if detail:
        H.append(f'<div class="ke-detail">{detail}</div>')
    H.append('</div>')

H.append('</div>')

# 天盘五干逻辑
H.append('<div style="text-align:center;font-size:11px;color:#666;margin-bottom:6px">')
H.append(f'天盘五干: {v("D",4)} {v("E",4)} | {v("H",4)} {v("I",4)} | {v("L",4)} {v("M",4)} | {v("P",4)} {v("Q",4)}')
H.append('</div>')

# 五行
H.append('<div class="wx">')
H.append(f'<span class="wx-huo">日干: {v("D",4)}</span>')
H.append('<span class="wx-huo">旺: 火旺</span><span class="wx-tu">相: 土相</span>')
H.append('<span class="wx-mu">休: 木休</span><span class="wx-shui">囚: 水囚</span>')
H.append('<span class="wx-jin">死: 金死</span>')
H.append('</div>')

# 神煞详细
H.append('<div class="shensha-section">')
H.append('<div class="shensha-title">神煞明细</div>')
H.append('<div class="tag-row">')
for r in [1,2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19]:
    for col in ['U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF']:
        val = v(col, r)
        label = v(col, 1)
        if val and label and val != label and val not in ('日德','大耗','将星','月德','天赦','文昌','成神','天喜','生气','天德','孤辰','劫煞','支德','寡宿','丧车','岁德','喝散','天解','地解','解神','旬首','旬奇','日解','出行','天马','日马','月马','年马','月德'):
            H.append(f'<span class="tag"><span class="tl">{label}</span> <span class="tv">{val}</span></span>')

H.append('</div></div>')
H.append(f'<div style="text-align:center;color:#bbb;font-size:10px;margin-top:10px;padding:8px 0">天星2.0 · 霎哈嘉瑜伽修习者</div>')
H.append('</div></body></html>')

out = 'G:/WorkBuddy/码农龙虾/tianxing-daliuren/index.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(H))
print(f'Generated: {out} ({sum(len(h) for h in H)} bytes)')
