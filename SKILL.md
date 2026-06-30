---
name: brand-pptx
description: "テンプレート PPTX と theme.json から、自社ブランド準拠の PowerPoint を生成・編集する。テンプレートを複製しレイアウトを選択 → プレースホルダーを差し替え、KPI カード・進捗バー・ステップフロー・テーブルなどの視覚パーツを自動配置する。Triggers: PPT, PPTX, PowerPoint, presentation, slides, deck, スライド, プレゼン, 資料作成"
license: MIT
metadata:
  version: "1.0.0"
  based_on: "MiniMax-AI/skills@pptx-generator (MIT License)"
---

# brand-pptx — テンプレート準拠の PowerPoint 生成スキル

## 概要

**テンプレート PPTX を複製し、レイアウトを選んで、プレースホルダーを差し替える**だけで、
そのテンプレートのデザイン（配色・装飾・ロゴ）に沿ったスライドを生成する。

このスキルの肝は **「見た目をプロンプトで指示しない」** こと。
配色・装飾・レイアウトはすべて **テンプレート側** に持たせ、生成コードは中身を流し込むだけ。
だから誰が・何度頼んでも、テンプレートから離れたスライドにはならない。

- **同梱のクリーンな既定テンプレート（`assets/template.pptx`, 16:9）** ですぐに動く。
- **自社テンプレート（.pptx）に差し替えれば**、そのブランドをそのまま再現できる。差し替え手順は下記「初回セットアップ」。

> **なぜテンプレートを使うのか**
> 「青系でモダンに」のようなテキスト指示だけでブランドを固定するのは現実的に不可能で、毎回ブレる。
> 配色・装飾・レイアウトを *PPTX というデータ* として固定し、そこから離れられないようにするのがこのスキルの設計思想。

---

## 初回セットアップ（最初に一度だけ確認する）

資料を作り始める前に、ユーザーに **一度だけ** 確認する。難しい設定は不要。

> **「自社のテンプレート（.pptx）はありますか？**
> **あれば、それに合わせて生成します。無ければ同梱のクリーンなテンプレートで作ります。」**

### 自社テンプレートを渡された場合（スキルが自動で設定する）

ユーザー設定は **スキルの外**（スキル更新で上書きされない場所）に保存する。既定の保存先は **`~/.config/brand-pptx/`**。

> **勝手に作らない**: このフォルダは「ユーザーが自社テンプレートを渡した＝カスタマイズに同意した」場合にのみ作成する。通常利用では何も作成・変更しない。作成前に **保存先をユーザーに一言伝え**、別の場所がよければ環境変数 `$BRAND_PPTX_HOME` で指定してもらう。

1. 保存先（既定 `~/.config/brand-pptx/`、または `$BRAND_PPTX_HOME`）をユーザーに伝えてから作成し、テンプレートをそこへコピーする（例: `~/.config/brand-pptx/company.pptx`）。
2. `python3 tools/inspect_template.py ~/.config/brand-pptx/company.pptx` でレイアウト一覧（index・名前・プレースホルダー）を取得する。
3. `~/.config/brand-pptx/theme.json` を作成する（スキル同梱の `theme.json` をコピーして編集するとよい）。`template.path` を自社テンプレート名（例: `"company.pptx"`）にし、`layoutMap` を自社レイアウトの index に合わせる。判断の目安:
   - **cover**: タイトル＋サブタイトルのプレースホルダーがある表紙レイアウト
   - **section**: 章扉（背景が濃い/ブランド色のことが多い）
   - **content**: タイトル＋本文（body）の標準レイアウト
   - **contentVisual**: タイトルのみ（本文 PH なし）＝図形を自由配置できるレイアウト
   - **ending**: クロージング向けレイアウト
4. 必要なら `~/.config/brand-pptx/theme.json` の `colors`（KPI カード等の描画色）も自社ブランドに合わせて更新する。
5. `python3 tools/make_sample.py` でサンプルを出し、崩れていないか確認する。

> 以降スキルは `~/.config/brand-pptx/theme.json` を **優先して**読む（無ければ同梱の既定）。だからスキル本体を更新しても自社設定は保持される。

### 無い場合

同梱の `assets/template.pptx` のまま進める。**設定は何も要らない。**

### テンプレートに合うレイアウトが無い役割があるとき（任意・opt-in）

自社テンプレートに、ある役割（例: セクション区切り）に向くレイアウトが無い場合:

1. その役割を `theme.json` の `layoutMap` で `null` にする（例: `"section": null`）。
2. `theme.json` の `setup.drawMissingRoles` を `true` にする（既定は `false`）。

すると、その役割のスライドだけ **theme の色でスライドに直接描画**して補う。
**テンプレート（マスター/レイアウト）自体は一切変更しない**ので、元のテンプレートは安全。
`false` のままなら、`null` の役割は白紙レイアウト + タイトルのみの簡素な出力になる。

---

## `theme.json` の役割

> **読み込み場所**: `~/.config/brand-pptx/theme.json` があればそれを優先（`$BRAND_PPTX_HOME` でも上書き可）。無ければスキル同梱の既定。ユーザー設定をスキル外に置けるので、スキル更新で消えない。

| キー | 役割 |
|------|------|
| `template.path` | 使うテンプレート（既定: `assets/template.pptx` / 自社: `assets/<your>.pptx`） |
| `layoutMap` | 役割（cover/section/content/contentVisual/ending）→ テンプレート内レイアウト index |
| `colors` | **コードが描く視覚パーツ**（KPI カード・進捗バー・ステップ・テーブル）の色。`#` を付けない 6 桁 hex |
| `fonts` / `typeScale` | プレースホルダーや視覚パーツに適用するフォント・サイズ |
| `setup.drawMissingRoles` | （任意・既定 `false`）`layoutMap` が `null` の役割を theme 色でスライドに合成描画する。テンプレート自体は変更しない |

> 表紙やセクションの**背景・装飾はテンプレートが持つ**ので、`colors` は主に「コードが描く部分」と「同梱既定テンプレートの再生成」に効く。
> 同梱テンプレートは `python3 tools/build_template.py` が `theme.json` の色から生成する（色を変えて再実行すると既定テンプレートの配色も変わる）。

---

## クイックリファレンス

| タスク | 方法 |
|--------|------|
| テンプレートのレイアウト確認 | `python3 tools/inspect_template.py [PPTX]` |
| 新規作成 | **テンプレートを複製 → `layoutMap` でレイアウト選択 → プレースホルダー差し替え** |
| サンプル生成（動作確認） | `python3 tools/make_sample.py` → `examples/sample.pptx` |
| 既定テンプレートの再生成 | `python3 tools/build_template.py` → `assets/template.pptx` |
| 既存ファイルの読み取り | `python3 -m markitdown deck.pptx` |
| 既存ファイルの編集 | [editing.md](references/editing.md) の XML 編集ワークフロー |
| テンプレートにないカスタムスライド | PptxGenJS で補助生成 → [pptxgenjs.md](references/pptxgenjs.md) |

---

## 新規作成ワークフロー

### Step 0: テーマとテンプレートの読み込み

```python
import json, os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER

def find_skill_dir():
    cands = []
    pr = os.environ.get("CLAUDE_PLUGIN_ROOT")
    if pr:
        cands.append(Path(pr))
    cands += [
        Path.home() / ".claude" / "skills" / "brand-pptx",
        Path.home() / ".agents" / "skills" / "brand-pptx",
        Path.home() / ".config" / "skills" / "brand-pptx",
        Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd(),
        Path.cwd(),
    ]
    for c in cands:
        if (c / "theme.json").exists():
            return c
    return Path.cwd()

SKILL_DIR = find_skill_dir()

def find_config_dir():
    # ユーザー設定（theme.json/自社テンプレ）は「スキル更新で消えない場所」を優先する。
    # 優先: $BRAND_PPTX_HOME → ~/.config/brand-pptx → （無ければ）スキル同梱の既定。
    for c in [os.environ.get("BRAND_PPTX_HOME"), Path.home() / ".config" / "brand-pptx"]:
        if c and (Path(c) / "theme.json").exists():
            return Path(c)
    return SKILL_DIR

CONFIG_DIR = find_config_dir()
theme = json.loads((CONFIG_DIR / "theme.json").read_text(encoding="utf-8"))
C  = {k: RGBColor.from_string(v) for k, v in theme["colors"].items()}
F  = theme["fonts"]
TS = theme["typeScale"]
LM = theme["layoutMap"]

# テンプレート解決: config_dir（ユーザー設定）→ skill_dir → 同梱既定 → 内蔵 の順
def resolve_template():
    tpath = (theme.get("template") or {}).get("path")
    if tpath:
        p = Path(tpath)
        if p.is_absolute():
            return p if p.exists() else None
        for base in (CONFIG_DIR, SKILL_DIR):
            if (base / tpath).exists():
                return base / tpath
    default = SKILL_DIR / "assets" / "template.pptx"
    return default if default.exists() else None

tpl = resolve_template()
prs = Presentation(str(tpl)) if tpl else Presentation()
print("config:", CONFIG_DIR, "| template:", tpl or "(builtin)")
```

### Step 1: 要件確認

トピック、対象者、目的、スライド枚数を確認する（[初回セットアップ](#初回セットアップ最初に一度だけ確認する)も忘れずに）。

### Step 2: スライド構成の計画

各スライドに **役割**（`layoutMap` のキー）を割り当てる。例:
```
Slide 1: cover          (タイトル + サブタイトル/日付)
Slide 2: section        (セクション区切り)
Slide 3-6: content / contentVisual
Slide 7: ending         (クロージング)
```

### Step 3: 生成（プレースホルダーを埋める）

**背景・装飾はテンプレートのレイアウトが持つ**。ここでは文字を流し込むだけ。
タイトル色は役割で決める（濃い背景の section / ending は白、白背景の cover / content は `dark`）。

```python
def delete_all_slides(prs):
    rels = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    while len(prs.slides._sldIdLst):
        rId = prs.slides._sldIdLst[0].attrib[rels]
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

def _ph(slide, idx):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None

def _title_ph(slide):
    """タイトルPH。idx=0 を優先し、無ければ型でフォールバック（自社テンプレ対応）。"""
    p0 = _ph(slide, 0)
    if p0 is not None:
        return p0
    for ph in slide.placeholders:
        if ph.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            return ph
    return None

def _style_text(tf, size, color, bold=False, align=PP_ALIGN.LEFT, font=None, anchor=None):
    if anchor is not None:
        tf.vertical_anchor = anchor
    for p in tf.paragraphs:
        p.alignment = align
        for r in (p.runs or [p.add_run()]):
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            if font:
                r.font.name = font

delete_all_slides(prs)

# --- カバー（白背景・濃い文字）---
slide = prs.slides.add_slide(prs.slide_layouts[LM["cover"]])
t = _title_ph(slide)
if t:
    t.text = "プレゼンテーションタイトル"
    _style_text(t.text_frame, TS["coverTitle"], C["dark"], bold=True,
                font=F["headingCJK"], anchor=MSO_ANCHOR.MIDDLE)
sub = _ph(slide, 1)
if sub:
    sub.text = "2026.01.01"
    _style_text(sub.text_frame, TS["coverSubtitle"], C["gray"], font=F["bodyCJK"])

# --- セクション区切り（濃い背景・白文字）---
slide = prs.slides.add_slide(prs.slide_layouts[LM["section"]])
t = _title_ph(slide)
if t:
    t.text = "セクションタイトル"
    _style_text(t.text_frame, TS["sectionTitle"], C["white"], bold=True,
                font=F["headingCJK"], anchor=MSO_ANCHOR.MIDDLE)

# --- コンテンツ（図形中心は contentVisual を使う）---
slide = prs.slides.add_slide(prs.slide_layouts[LM["contentVisual"]])
t = _title_ph(slide)
if t:
    t.text = "コンテンツタイトル"
    _style_text(t.text_frame, TS["title"], C["dark"], bold=True, font=F["headingCJK"])
# → ここに Step 4 のヘルパーで KPI カード等を配置

# --- エンディング（濃い背景・白文字・中央）---
slide = prs.slides.add_slide(prs.slide_layouts[LM["ending"]])
t = _title_ph(slide)
if t:
    t.text = "ご清聴ありがとうございました"
    _style_text(t.text_frame, TS["sectionTitle"], C["white"], bold=True,
                align=PP_ALIGN.CENTER, font=F["headingCJK"], anchor=MSO_ANCHOR.MIDDLE)

os.makedirs("output", exist_ok=True)
prs.save("output/presentation.pptx")
print("saved: output/presentation.pptx")
```

> `tools/brandkit.py` に同じロジックの `Brand` クラス（`add_cover` / `add_section` / `add_content` / `add_ending` と各視覚パーツ）がある。スクリプトから使うとワークフローが短くなる。

### Step 4: 図形・ビジュアルの追加

コンテンツスライドは **テキストだけにしない**。数値は KPI カード、進捗はバー、手順はステップフローで表現する。
色は `theme.json`（`C[...]`）から取り、**パレット外の色を直接書かない。**

**重要なルール:**
- テキストフレームには必ずマージンを設定する（はみ出し防止）
- 図形内テキストの中央揃えは `word_wrap=True` + `alignment=CENTER` + `vertical_anchor=MIDDLE` を 3 点セットで
- 要素の重なり防止: 前要素の bottom (y+h) に gap を足して次の y を決める。座標をハードコードで並べない
- スライドの安全領域に収める（既定テンプレートは 16:9 = 13.333"×7.5"。左右マージンは `style.pageMarginIn`）
- 角丸は `theme.json` の `style.cornerRadiusIn` に合わせる: ROUNDED_RECTANGLE は `shape.adjustments[0] = min(0.5, int(Inches(style["cornerRadiusIn"])) / min(w, h))`（`tools/brandkit.py` の `round_corners()` 参照）
- **バランス（中央揃え）を必ず取る**: カードやステップを横に並べるときは **コンテンツ幅いっぱいに等幅**で配置し、左右の余白を均等にする（左に寄せない。`brandkit.py` の `add_kpi_row()` 参照）。矢印などのコネクタは要素と要素の **「間の中央」** に置く。タイトルはプレースホルダー内で `vertical_anchor=MIDDLE`

```python
def _set_margins(tf, l=Inches(0.1), r=Inches(0.1), t=Inches(0.05), b=Inches(0.05)):
    tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom = l, r, t, b
    tf.word_wrap = True

def add_kpi_card(slide, x, y, w, h, value, label):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = C["bg"]; card.line.fill.background()
    val_h = int(h * 0.55)
    b1 = slide.shapes.add_textbox(x, y, w, val_h); tf = b1.text_frame
    _set_margins(tf); tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.paragraphs[0].text = str(value)
    tf.paragraphs[0].font.size = Pt(TS["kpiValue"]); tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = C["accent"]; tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    b2 = slide.shapes.add_textbox(x, y + val_h, w, h - val_h); tf2 = b2.text_frame
    _set_margins(tf2); tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf2.paragraphs[0].text = label
    tf2.paragraphs[0].font.size = Pt(TS["caption"]); tf2.paragraphs[0].font.color.rgb = C["gray"]
    tf2.paragraphs[0].alignment = PP_ALIGN.CENTER

def add_progress_bar(slide, x, y, w, h, progress):
    """progress: 0.0 ~ 1.0"""
    bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    bg.fill.solid(); bg.fill.fore_color.rgb = C["bg"]; bg.line.fill.background()
    bw = int(w * max(0.0, min(1.0, progress)))
    if bw > 0:
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, bw, h)
        bar.fill.solid(); bar.fill.fore_color.rgb = C["accent"]; bar.line.fill.background()

# KPI カード・進捗バー・ステップフロー・テーブルの完全版は tools/brandkit.py を参照
```

### Step 5: QA（必須）

[QA プロセス](references/pitfalls.md#qa-process)。最低 1 回は「生成 → markitdown で抽出 → 修正」を回す。

---

## 既存ファイルの読み取り

```bash
python3 -m markitdown deck.pptx
```

## レファレンスファイル

| ファイル | 内容 |
|----------|------|
| [design-system.md](references/design-system.md) | `theme.json` の設計指針・配色/タイポ/スペーシングの考え方 |
| [slide-types.md](references/slide-types.md) | スライドタイプとレイアウトパターン（`layoutMap` ロールとの対応表つき） |
| [editing.md](references/editing.md) | 既存 PPTX の XML 編集ワークフロー |
| [pitfalls.md](references/pitfalls.md) | QA プロセス、よくあるミス |
| [pptxgenjs.md](references/pptxgenjs.md) | PptxGenJS API リファレンス（テンプレート外のカスタムスライド用） |

## 依存パッケージ

```bash
pip install python-pptx "markitdown[pptx]"
npm install pptxgenjs   # テンプレート外のカスタムスライド用（任意）
```
