"""Render each PPTX slide as a PNG preview using Pillow."""
from PIL import Image, ImageDraw, ImageFont
import os, math

OUT = "/home/user/contact-form/preview"
os.makedirs(OUT, exist_ok=True)

W, H = 1333, 750

# ── Colors ──────────────────────────────────────────────────────
NAVY   = (26,  46,  72)
BLUE   = (40,  96, 168)
LBLUE  = (74, 144, 200)
PALE   = (238, 245, 252)
ACCENT = (232, 160,  32)
GRAY   = (106, 122, 138)
WHITE  = (255, 255, 255)
LGRAY  = (200, 220, 234)
SILVER = (220, 232, 240)
BG_DARK= (13,  28,  46)

# ── Font loading ─────────────────────────────────────────────────
def get_font(size=14, bold=False):
    paths = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for p in paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except:
                continue
    return ImageFont.load_default()

# Pre-load common sizes
F = {s: get_font(s) for s in [7,8,9,10,11,12,13,14,16,18,20,22,24,28,32,36,44]}
FB= {s: get_font(s, bold=True) for s in [7,8,9,10,11,12,13,14,16,18,20,22,24,28,32,36,44]}

# ── Drawing helpers ──────────────────────────────────────────────
def scale(v):
    """Scale from ~100px unit to slide width"""
    return int(v)

def rect(draw, x, y, w, h, fill=None, outline=None, radius=8):
    if radius:
        draw.rounded_rectangle([x, y, x+w, y+h], radius=radius,
                                fill=fill, outline=outline, width=2 if outline else 0)
    else:
        draw.rectangle([x, y, x+w, y+h], fill=fill, outline=outline,
                       width=2 if outline else 0)

def text(draw, x, y, t, font_size=12, color=NAVY, bold=False,
         align="left", max_w=None):
    f = FB[font_size] if bold else F[font_size]
    if align == "center" and max_w:
        bb = f.getbbox(t)
        tw = bb[2] - bb[0]
        x = x + (max_w - tw) // 2
    elif align == "right" and max_w:
        bb = f.getbbox(t)
        tw = bb[2] - bb[0]
        x = x + max_w - tw
    draw.text((x, y), t, font=f, fill=color)

def multiline(draw, x, y, t, font_size=11, color=NAVY, bold=False,
              line_h=None, max_w=None, align="left"):
    f = FB[font_size] if bold else F[font_size]
    lh = line_h or (font_size + 4)
    lines = t.split("\n")
    for i, line in enumerate(lines):
        lx = x
        if align == "center" and max_w:
            bb = f.getbbox(line)
            lx = x + (max_w - (bb[2]-bb[0])) // 2
        draw.text((lx, y + i*lh), line, font=f, fill=color)

def header(draw, page_tag, title, right_text=""):
    rect(draw, 0, 0, W, 72, fill=NAVY, radius=0)
    text(draw, 28, 10, page_tag, 9, LBLUE)
    text(draw, 28, 28, title, 18, WHITE, bold=True)
    if right_text:
        f = F[9]
        bb = f.getbbox(right_text)
        tw = bb[2]-bb[0]
        draw.text((W-28-tw, 35), right_text, font=f, fill=(160,180,200))

def card(draw, x, y, w, h, fill=PALE, outline=LGRAY, r=10):
    rect(draw, x, y, w, h, fill=fill, outline=outline, radius=r)

def accent_bar(draw, x, y, w, color=BLUE):
    rect(draw, x, y, w, 5, fill=color, radius=0)

def arrow_h(draw, x, y, length, label="", sublabel="", color=BLUE):
    draw.line([(x, y), (x+length-10, y)], fill=color, width=2)
    draw.polygon([(x+length-10, y-5), (x+length, y), (x+length-10, y+5)], fill=color)
    if label:
        f = F[8]
        bb = f.getbbox(label)
        lx = x + (length-(bb[2]-bb[0]))//2
        draw.text((lx, y-18), label, font=f, fill=GRAY)
    if sublabel:
        f = F[8]
        bb = f.getbbox(sublabel)
        lx = x + (length-(bb[2]-bb[0]))//2
        draw.text((lx, y+6), sublabel, font=f, fill=color)

def ph_bar(draw, x, y, w, h=12, color=SILVER):
    rect(draw, x, y, w, h, fill=color, radius=4)

def new_slide():
    img = Image.new("RGB", (W, H), WHITE)
    draw = ImageDraw.Draw(img)
    return img, draw

# ═══════════════════════════════════════════════════════════════
# SLIDE 1 – 会社概要
# ═══════════════════════════════════════════════════════════════
def slide1():
    img, draw = new_slide()
    header(draw, "01 / Company Overview", "会  社  概  要",
           "SHIKI（親会社）× 日本住戸（グループ会社）")

    # banner
    rect(draw, 26, 76, W-52, 36, fill=NAVY, radius=8)
    t1 = "SHIKI は 日本住戸 の"
    t2 = " 親会社（出資元） "
    t3 = "であり、同一グループとして協業します"
    bx = 36
    by = 84
    text(draw, bx, by, t1, 9, (160,190,210))
    f = FB[9]; bb = F[9].getbbox(t1); bx2 = bx + bb[2]-bb[0]
    draw.text((bx2, by), t2, font=FB[9], fill=LBLUE)
    bb2 = FB[9].getbbox(t2); bx3 = bx2+bb2[2]-bb2[0]
    draw.text((bx3, by), t3, font=F[9], fill=(160,190,210))

    cw = (W-26*2-20)//2
    Lx, Rx = 26, 26+cw+20
    cy = 118
    ch = H-cy-16

    # SHIKI card
    card(draw, Lx, cy, cw, ch)
    accent_bar(draw, Lx, cy, cw, BLUE)
    tx = Lx+18; ty = cy+14
    lw = cw-36
    text(draw, tx, ty, "SHIKI", 22, BLUE, bold=True); ty += 30
    rect(draw, tx, ty, 120, 22, fill=BLUE, radius=11)
    text(draw, tx+10, ty+4, "親会社・出資元", 8, WHITE, bold=True); ty += 32
    rows = [("会社名","株式会社 SHIKI"),("所在地","●●●●●●●●●●"),
            ("設　立","●●●●年●月"),("代　表","●●　●●"),
            ("資本金","●●●万円"),("事業内容","不動産テック・プラットフォーム事業")]
    for lbl, val in rows:
        text(draw, tx, ty, lbl, 8, GRAY); text(draw, tx+82, ty, val, 10, NAVY); ty+=24

    # 出資矢印
    mx = Lx+cw+4; my = cy+ch//2
    rect(draw, mx, my-1, 12, 3, fill=BLUE, radius=0)
    text(draw, mx+1, my-18, "出資", 8, BLUE)
    draw.polygon([(mx+6, my+4),(mx+12, my+12),(mx, my+12)], fill=BLUE)

    # 日本住戸 card
    card(draw, Rx, cy, cw, ch)
    accent_bar(draw, Rx, cy, cw, NAVY)
    tx = Rx+18; ty = cy+14
    text(draw, tx, ty, "日本住戸", 22, NAVY, bold=True); ty += 30
    rect(draw, tx, ty, 150, 22, fill=NAVY, radius=11)
    text(draw, tx+10, ty+4, "グループ会社（子会社）", 8, WHITE, bold=True); ty += 32
    rows2=[("会社名","●●●●株式会社"),("所在地","●●●●●●●●●●"),
           ("設　立","●●●●年●月"),("代　表","●●　●●"),
           ("資本金","●●●万円"),("事業内容","不動産管理・賃貸仲介事業")]
    for lbl, val in rows2:
        text(draw, tx, ty, lbl, 8, GRAY); text(draw, tx+82, ty, val, 10, NAVY); ty+=24

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 2 – ビジネスモデル
# ═══════════════════════════════════════════════════════════════
def slide2():
    img, draw = new_slide()
    header(draw, "02 / Business Model", "ビ ジ ネ ス モ デ ル", "SHIKIモデルを踏襲")

    # desc bar
    card(draw, 26, 78, W-52, 44, fill=PALE, outline=LGRAY)
    rect(draw, 26, 78, 6, 44, fill=BLUE, radius=0)
    text(draw, 42, 86,
         "SHIKIが確立した月次課金モデルを基盤に、物件オーナーへのバリュー提供と仲介会社への収益シェアを組み合わせた三方よしのビジネス設計です。",
         9, NAVY)

    # flow boxes
    boxes = [
        ("🏢", "物件オーナー",  "管理コスト削減\n空室リスク低減", False),
        ("⚡", "SHIKI×日本住戸\nPlatform",       "プラットフォーム\n運営・管理", True),
        ("🏪", "仲 介 会 社",  "手数料収入\nストック収益",  False),
        ("👤", "入 居 者",     "月500円〜\n付帯サービス享受", False),
    ]
    bw, bh = 200, 160
    gap = 52
    total = len(boxes)*bw + (len(boxes)-1)*gap
    sx = (W-total)//2
    by = 140

    for i,(icon,title,sub,primary) in enumerate(boxes):
        bx = sx + i*(bw+gap)
        bg = NAVY if primary else PALE
        bc = BLUE if primary else LGRAY
        tc = WHITE if primary else NAVY
        sc = (170,210,240) if primary else GRAY
        card(draw, bx, by, bw, bh, fill=bg, outline=bc)
        text(draw, bx, by+10, icon, 24, tc, align="center", max_w=bw)
        multiline(draw, bx, by+46, title, 11, tc, bold=True, line_h=16, max_w=bw, align="center")
        multiline(draw, bx, by+86, sub, 9, sc, line_h=15, max_w=bw, align="center")

        if i < len(boxes)-1:
            ax = bx+bw+6
            ay = by+bh//2
            lbl = ["委託","紹介","マッチング"][i]
            sub_lbl = ["収益還元","手数料","月500円〜"][i]
            arrow_h(draw, ax, ay, gap-12, lbl, sub_lbl)

    # feature cards
    feats=[("FEATURE 01","月次サブスクリプション","入居期間中継続的な収益を確保"),
           ("FEATURE 02","複数収益源の統合","電気・コンテンツ・ネットを一本化"),
           ("FEATURE 03","スケーラブル設計","戸数増加に応じたストック積み上げ")]
    fw=(W-52-20*2)//3; fy=H-105
    for j,(tag,ttl,sub) in enumerate(feats):
        fx=26+j*(fw+20)
        card(draw, fx, fy, fw, 95, fill=PALE, outline=LGRAY)
        text(draw, fx+12, fy+8, tag, 7, GRAY)
        text(draw, fx+12, fy+24, ttl, 11, NAVY, bold=True)
        text(draw, fx+12, fy+46, sub, 9, GRAY)

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 3 – 仲介ネットワーク
# ═══════════════════════════════════════════════════════════════
def slide3():
    img, draw = new_slide()
    header(draw, "03 / Broker Network", "仲 介 ネ ッ ト ワ ー ク", "SHIKIネットワークを踏襲")

    cw=(W-52-20)//2; Lx,Rx=26,26+cw+20; top=78

    # Hub circle
    hcx=Lx+cw//2; hcy=top+130; hr=80
    draw.ellipse([hcx-hr,hcy-hr,hcx+hr,hcy+hr], fill=NAVY, outline=BLUE, width=2)
    text(draw, hcx, hcy-38, "🔗", 22, WHITE, align="center", max_w=0)
    text(draw, hcx-28, hcy-10, "SHIKI", 13, WHITE, bold=True)
    text(draw, hcx-30, hcy+10, "Platform", 9, LBLUE)
    text(draw, hcx-22, hcy+28, "全国対応", 8, (170,200,230))

    # Stats grid
    stats=[("●●社","提携仲介会社数"),("●●都道府県","カバーエリア"),
           ("●●万戸","管理物件数"),("●●年","構築年数")]
    sw=(cw-12)//2; sy_base=top+230
    for k,(num,lbl) in enumerate(stats):
        sx=Lx+(k%2)*(sw+12); sy=sy_base+(k//2)*72
        card(draw, sx, sy, sw, 64, fill=PALE, outline=LGRAY)
        text(draw, sx, sy+8, num, 16, BLUE, bold=True, align="center", max_w=sw)
        text(draw, sx, sy+40, lbl, 8, GRAY, align="center", max_w=sw)

    # Spokes
    spokes=[("🏪","大手仲介チェーン","全国展開の大手不動産会社と直接提携"),
            ("🏘️","地域密着型仲介","地元に根付いた中小仲介会社ネットワーク"),
            ("💻","オンライン仲介","デジタル仲介プラットフォームとAPI連携"),
            ("🤝","管理会社直接提携","物件管理会社を通じた一括導入スキーム")]
    sh=68; sgap=10; sy0=top+(H-top-16-len(spokes)*(sh+sgap))//2
    for k,(icon,ttl,sub) in enumerate(spokes):
        sy=sy0+k*(sh+sgap)
        card(draw, Rx, sy, cw, sh, fill=PALE, outline=LGRAY)
        text(draw, Rx+12, sy+16, icon, 18)
        text(draw, Rx+52, sy+10, ttl, 11, NAVY, bold=True)
        text(draw, Rx+52, sy+30, sub, 9, GRAY)

    # note
    ny=sy0+len(spokes)*(sh+sgap)+10
    rect(draw, Rx, ny, cw, 48, fill=NAVY, radius=8)
    multiline(draw, Rx+12, ny+8,
              "SHIKIのネットワーク資産をそのまま活用することで、\n日本住戸は初期コストゼロで全国展開が可能です。",
              9, (220,235,255), line_h=18)

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 4 – 商品概要
# ═══════════════════════════════════════════════════════════════
def slide4():
    img, draw = new_slide()
    header(draw, "04 / Product", "商  品  概  要", "料金プラン")

    cw=(W-52-20)//2; Lx,Rx=26,26+cw+20; top=78; ch=H-top-16

    def pcol(cx, ac, label, price, unit, desc, feats):
        card(draw, cx, top, cw, ch, fill=PALE, outline=LGRAY)
        rect(draw, cx, top, cw, 5, fill=ac, radius=0)
        tx=cx+22; ty=top+18; lw=cw-44
        text(draw, tx, ty, label, 8, GRAY); ty+=22
        text(draw, tx, ty, price, 44, ac, bold=True); ty+=58
        text(draw, tx, ty, unit, 9, GRAY); ty+=22
        rect(draw, tx, ty, lw, 1, fill=LGRAY); ty+=12
        multiline(draw, tx, ty, desc, 9, NAVY, line_h=16); ty+=46
        for f in feats:
            text(draw, tx, ty, "✓ "+f, 10, NAVY); ty+=24

    pcol(Lx, BLUE, "月額基本プラン", "¥500", "/ 月 · 入居者1名あたり",
         "入居者が毎月負担する基本利用料。\n付帯サービスへのアクセス権が含まれます。",
         ["電気・コンテンツ・ネットが利用可能","入居期間中継続課金","初期費用なし"])
    pcol(Rx, ACCENT, "仲介手数料シェア", "  ●%", "収益に対する仲介会社への還元率",
         "月額収益の一定割合を仲介会社へ自動還元。\n入居者在籍中はストック収益が積み上がります。",
         ["自動精算・振込","リアルタイムダッシュボード管理","戸数に応じた段階レート"])

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 5 – 付帯サービス
# ═══════════════════════════════════════════════════════════════
def slide5():
    img, draw = new_slide()
    header(draw, "05 / Add-on Services", "付 帯 サ ー ビ ス", "3つのバリュー")

    n=3; gap=16; cw=(W-52-gap*(n-1))//n; top=78; ch=H-top-16
    services=[
        ("⚡","電　　　気","ENERGY",(255,243,224),(232,160,32),
         "入居者向けの電力供給サービス。\n割安な電気料金プランと使用量の\n可視化・節電サポートを提供します。",
         "入居者メリット","●●●●●●●●"),
        ("🎬","コンテンツ","CONTENTS",(243,229,245),(144,96,184),
         "動画・音楽・電子書籍などの\nデジタルコンテンツを月額サービスに\n含めて提供します。",
         "提供コンテンツ","●●●●●●●●"),
        ("📡","インターネット","NETWORK",(227,242,253),(32,128,192),
         "物件全体への高速インターネット\n回線を一括導入。入居者は追加\n契約不要で即日利用可能です。",
         "回線スペック","●●●●●●●●"),
    ]
    for i,(icon,name,tag,ibg,ifg,desc,blbl,bval) in enumerate(services):
        cx=26+i*(cw+gap)
        card(draw, cx, top, cw, ch, fill=PALE, outline=LGRAY)
        tx=cx+18; ty=top+18; lw=cw-36
        # icon box
        rect(draw, tx, ty, 56, 56, fill=ibg, radius=12)
        text(draw, tx, ty+10, icon, 24, align="center", max_w=56)
        ty+=72
        text(draw, tx, ty, name, 14, NAVY, bold=True); ty+=24
        rect(draw, tx, ty, 88, 20, fill=BLUE, radius=10)
        text(draw, tx+8, ty+4, tag, 8, WHITE, bold=True); ty+=32
        multiline(draw, tx, ty, desc, 9, GRAY, line_h=16); ty+=60
        # benefit box
        card(draw, tx, ty, lw, 58, fill=WHITE, outline=LGRAY, r=8)
        text(draw, tx+10, ty+7, blbl, 7, GRAY)
        text(draw, tx+10, ty+24, bval, 11, BLUE, bold=True)

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 6 – 提案
# ═══════════════════════════════════════════════════════════════
def slide6():
    img, draw = new_slide()
    header(draw, "06 / Proposal", "提　　　　案", "スキーム図")

    note_h=50; note_y=H-note_h-14
    rect(draw, 26, note_y, W-52, note_h, fill=NAVY, radius=8)
    multiline(draw, 40, note_y+8,
              "提案要旨：SHIKIが構築したプラットフォームと仲介ネットワークを活用し、日本住戸の物件・顧客基盤を組み合わせることで、",
              9, (220,235,255), line_h=16)
    text(draw, 40, note_y+26,
         "入居者への付加価値提供・仲介会社への継続収益・オーナーへのコスト削減を同時に実現します。",
         9, (220,235,255))

    # Scheme area
    diag_top=80; diag_h=note_y-diag_top-10
    card(draw, 26, diag_top, W-52, diag_h, fill=PALE, outline=LGRAY)

    bw=190; bh=155; gap=48
    n=4; total=n*bw+(n-1)*gap; sx=(W-total)//2; cy=diag_top+(diag_h-bh)//2

    nodes=[
        (False, "🏢", "物件オーナー",   "物件提供\n管理委託"),
        (True,  "⚡", "SHIKI×日本住戸\nPlatform","月次課金管理\nサービス統合"),
        (False, "🏪", "仲 介 会 社",   "入居者紹介\nストック収益"),
        (False, "👤", "入 居 者",      "月額課金\nサービス享受"),
    ]
    for i,(primary,icon,title,sub) in enumerate(nodes):
        bx=sx+i*(bw+gap)
        bg=NAVY if primary else PALE
        bc=BLUE if primary else LGRAY
        tc=WHITE if primary else NAVY
        sc=(170,210,240) if primary else GRAY
        card(draw, bx, cy, bw, bh, fill=bg, outline=bc)
        text(draw, bx, cy+8, icon, 22, tc, align="center", max_w=bw)
        multiline(draw, bx, cy+42, title, 10, tc, bold=True, line_h=15, max_w=bw, align="center")
        multiline(draw, bx, cy+80, sub, 9, sc, line_h=15, max_w=bw, align="center")

        if i < n-1:
            ax=bx+bw+5; ay=cy+bh//2
            lbls=["委託","紹介","マッチング"]; subs=["収益還元","手数料","月500円〜"]
            arrow_h(draw, ax, ay, gap-10, lbls[i], subs[i])

    # services row
    svc_y=cy+bh+18
    svc_items=[("⚡ 電気",(255,243,224),(232,160,32)),
               ("🎬 コンテンツ",(243,229,245),(144,96,184)),
               ("📡 ネット",(227,242,253),(32,128,192))]
    sw2=110; px=sx+(bw+gap)+(bw-sw2*3-20)//2
    for k,(lbl,bg2,fg2) in enumerate(svc_items):
        xi=px+k*(sw2+10)
        rect(draw, xi, svc_y, sw2, 28, fill=bg2, outline=fg2, radius=8)
        text(draw, xi, svc_y+7, lbl, 9, fg2, align="center", max_w=sw2)
        bx2=sx+(bw+gap); mid=bx2+bw//2
        draw.line([(xi+sw2//2, cy+bh),(xi+sw2//2, svc_y)], fill=GRAY, width=1)

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 7 – 手数料
# ═══════════════════════════════════════════════════════════════
def slide7():
    img, draw = new_slide()
    header(draw, "07 / Fee Structure", "手 数 料 イ メ ー ジ", "ストック収益モデル")

    cw=(W-52-20)//2; Lx,Rx=26,26+cw+20; top=78; ch=H-top-16

    # Left: fee breakdown
    card(draw, Lx, top, cw, ch, fill=PALE, outline=LGRAY)
    tx=Lx+18; ty=top+18; lw=cw-36
    text(draw, tx, ty, "手数料内訳イメージ", 9, GRAY, bold=True); ty+=28
    rows=[("入居者月額基本料","¥500 / 月"),("電気サービス収益","●% シェア"),
          ("コンテンツ収益","●% シェア"),("インターネット収益","●% シェア")]
    for name,val in rows:
        text(draw, tx, ty, name, 10, NAVY)
        vf=FB[12]; vbb=vf.getbbox(val)
        draw.text((tx+lw-(vbb[2]-vbb[0]), ty-2), val, font=vf, fill=BLUE)
        ty+=8; rect(draw, tx, ty, lw, 1, fill=LGRAY); ty+=14
    rect(draw, tx, ty, lw, 2, fill=BLUE); ty+=10
    text(draw, tx, ty, "仲介会社への還元率", 11, NAVY, bold=True)
    draw.text((tx+lw-50, ty-4), "●%", font=FB[22], fill=BLUE)
    ty+=44
    card(draw, tx, ty, lw, 52, fill=WHITE, outline=LGRAY, r=6)
    multiline(draw, tx+10, ty+8,
              "※ 入居者が在籍する限り毎月継続して受け取れる\n　 ストック型収益。戸数が増えるほど積み上がります。",
              8, GRAY, line_h=16)

    # Right: bar chart
    card(draw, Rx, top, cw, ch, fill=PALE, outline=LGRAY)
    tx2=Rx+18; ty2=top+18; lw2=cw-36
    text(draw, tx2, ty2, "ストック収益積み上がりイメージ", 9, GRAY, bold=True); ty2+=28

    chart_data=[1,2,3.2,4.5,6.2,8,9.8,11.5,13.5,15.2,17,19]
    chart_h=220; chart_y=ty2+8; max_v=max(chart_data)
    nb=len(chart_data); bw2=(lw2-4*(nb-1))//nb
    for j,v in enumerate(chart_data):
        bh_px=int(chart_h*(v/max_v))
        bx=tx2+j*(bw2+4); by_=chart_y+chart_h-bh_px
        # gradient: lighter top, darker bottom
        for row in range(bh_px):
            ratio=row/max(bh_px,1)
            r=int(74+(40-74)*ratio); g=int(144+(96-144)*ratio); b=int(200+(168-200)*ratio)
            draw.rectangle([bx, by_+row, bx+bw2, by_+row+1], fill=(r,g,b))
        draw.rounded_rectangle([bx,by_,bx+bw2,by_+bh_px], radius=3, fill=None, outline=None)

    rect(draw, tx2, chart_y+chart_h, lw2, 1, fill=LGRAY, radius=0)

    ty2=chart_y+chart_h+14
    text(draw, tx2, ty2, "※ 保有戸数・稼働率に応じた試算イメージ。実際の収益は条件により異なります。", 7, GRAY); ty2+=18

    sw3=(lw2-12)//2
    for k,(unit,val,col) in enumerate([("100戸×24ヶ月 目安","●●万円/月",BLUE),
                                        ("500戸×24ヶ月 目安","●●万円/月",ACCENT)]):
        sx3=tx2+k*(sw3+12)
        card(draw, sx3, ty2, sw3, 60, fill=WHITE, outline=LGRAY, r=6)
        text(draw, sx3, ty2+7, unit, 7, GRAY, align="center", max_w=sw3)
        text(draw, sx3, ty2+26, val, 13, col, bold=True, align="center", max_w=sw3)

    return img

# ═══════════════════════════════════════════════════════════════
# SLIDE 8 – 問い合わせ先
# ═══════════════════════════════════════════════════════════════
def slide8():
    img, draw = new_slide()
    header(draw, "08 / Contact", "問 い 合 わ せ 先")

    cw=(W-52-20)//2; Lx,Rx=26,26+cw+20; top=78
    footer_h=52; ch=H-top-footer_h-30

    def ccard(cx, brand, bc, rows):
        card(draw, cx, top, cw, ch, fill=PALE, outline=LGRAY)
        tx=cx+24; ty=top+20; lw=cw-48
        text(draw, tx, ty, brand, 20, bc, bold=True); ty+=32
        rect(draw, tx, ty, lw, 1, fill=LGRAY); ty+=12
        for lbl,val in rows:
            text(draw, tx, ty, lbl, 7, GRAY); ty+=16
            text(draw, tx, ty, val, 11, NAVY, bold=True); ty+=28

    ccard(Lx, "SHIKI", BLUE,
          [("担当部署","●●●●部"),("担当者名","●●　●●"),
           ("電話番号","0●-●●●●-●●●●"),("メールアドレス","●●●●@shiki.co.jp")])
    ccard(Rx, "日本住戸", NAVY,
          [("担当部署","●●●●部"),("担当者名","●●　●●"),
           ("電話番号","0●-●●●●-●●●●"),("メールアドレス","●●●●@nihon-juto.co.jp")])

    fy=H-footer_h-8
    rect(draw, 26, fy, W-52, footer_h, fill=NAVY, radius=8)
    text(draw, 0, fy+14,
         "お気軽にお問い合わせください　|　SHIKI（親会社）× 日本住戸（グループ会社）",
         10, (180,200,220), align="center", max_w=W)

    return img

# ═══════════════════════════════════════════════════════════════
# Render all slides
# ═══════════════════════════════════════════════════════════════
slides = [
    (slide1, "slide_01_会社概要"),
    (slide2, "slide_02_ビジネスモデル"),
    (slide3, "slide_03_仲介ネットワーク"),
    (slide4, "slide_04_商品概要"),
    (slide5, "slide_05_付帯サービス"),
    (slide6, "slide_06_提案"),
    (slide7, "slide_07_手数料"),
    (slide8, "slide_08_問い合わせ先"),
]

paths = []
for fn, name in slides:
    img = fn()
    path = f"{OUT}/{name}.png"
    img.save(path)
    paths.append(path)
    print(f"  ✓ {name}")

print(f"\nSaved {len(paths)} slide previews to {OUT}/")
