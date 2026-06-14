# スライドタイプ

すべてのスライドを**次の分類のいずれか**に当てはめる。
コード例は PptxGenJS（`theme` は `theme.json` から読み込んだ色オブジェクト）で示すが、同じレイアウトの考え方は python-pptx + テンプレートでも成り立つ。色は必ず `theme.*` から取り、直書きしない。

> **スライドタイプと `layoutMap` ロールの対応**
> 「スライドタイプ（見せ方の分類）」と `theme.json` の `layoutMap`（レイアウト割り当てのロール）は別物。下表で対応づける。
>
> | スライドタイプ | `layoutMap` ロール | 備考 |
> |----------------|--------------------|------|
> | カバー | `cover` | |
> | 目次（Index） | `content` | 専用ロールは無い。コンテンツレイアウトで作る |
> | セクション区切り | `section` | |
> | コンテンツ（テキスト中心） | `content` | |
> | コンテンツ（図形中心） | `contentVisual` | コンテンツページの図形中心サブタイプ |
> | まとめ / クロージング | `ending` | |

---

## 1. カバーページ

- **用途**: オープニング + トーン設定
- **内容**: タイトル、サブタイトル/発表者、日付、（任意で）ロゴ・装飾

### レイアウト例

```
|  タイトル                     |                    |
|  2026.01.01                  |  [装飾 / 余白]      |
|  [logo（任意）]               |                    |
```

### フォントサイズ

| 要素 | 推奨 (pt) |
|------|-----------|
| メインタイトル | 44 〜 60 (`coverTitle`) |
| サブタイトル / 日付 | 18 〜 24 (`coverSubtitle`) |
| 補足情報 | 14 〜 18 |

### コード例

```javascript
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.white };

  slide.addText("プレゼンテーションタイトル", {
    x: 0.5, y: 1.8, w: 5.5, h: 1.5,
    fontSize: 48, fontFace: "Noto Sans JP", color: theme.dark, bold: true
  });
  slide.addText("2026.01.01", {
    x: 0.5, y: 3.4, w: 5, h: 0.5,
    fontSize: 20, fontFace: "Arial", color: theme.gray
  });
  // ロゴを使う場合のみ（theme.logo.enabled）
  // slide.addImage({ path: theme.logo.lightPath, x: 0.5, y: 4.3, w: 2.0, h: 0.45 });
}
```

---

## 2. 目次（Index）

- **用途**: ナビゲーション + 期待値設定（3〜7 セクション）
- **内容**: 番号付きセクションリスト

### フォントサイズ

| 要素 | 推奨 (pt) |
|------|-----------|
| "index" タイトル | 36 〜 40 |
| セクション番号 + タイトル | 18 〜 22 |
| サブ項目 / 時間表示 | 14 〜 16 |

### コード例

```javascript
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.white };

  slide.addText("index", {
    x: 0.5, y: 0.3, w: 8, h: 0.8,
    fontSize: 36, fontFace: "Arial", color: theme.dark, bold: true
  });

  const items = ["1.　現状サマリー", "2.　主要トピック", "3.　次のアクション"];
  items.forEach((item, i) => {
    slide.addText(item, {
      x: 0.8, y: 1.3 + i * 0.55, w: 8, h: 0.5,
      fontSize: 20, fontFace: "Noto Sans JP", color: theme.dark
    });
  });
}
```

---

## 3. セクション区切り

- **用途**: セクション間の明確な区切り
- **内容**: セクション番号 + タイトル

### レイアウト例

```
| ████████████████████████████████████ |   ← 全面 theme.primary
| ████  Section 1                ████ |
| ████  セクションタイトル           ████ |
| ████  [logo（任意・反転色）]       ████ |
| ████████████████████████████████████ |
```

- 全面 `theme.primary` 背景、テキストは `theme.white`
- ロゴを使う場合は反転色（明るいロゴ）

### コード例

```javascript
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };

  slide.addText("Section 1", {
    x: 1.5, y: 2.0, w: 7, h: 0.5,
    fontSize: 18, fontFace: "Noto Sans JP", color: theme.white
  });
  slide.addText("セクションタイトル", {
    x: 1.5, y: 2.5, w: 7, h: 1,
    fontSize: 44, fontFace: "Noto Sans JP", color: theme.white, bold: true
  });
}
```

---

## 4. コンテンツページ

コンテンツに応じてサブタイプを選ぶ。各スライドは 1 つのサブタイプに属する。

### サブタイプ

- **テキスト** — 箇条書き・短い段落。**テキストだけにしない**。必ずアイコンや図形を 1 つ以上含める
- **混合メディア** — 2 カラム、または画像 + テキスト
- **データ可視化** — 図形ベースのバー/プログレス/リング + 要点
- **比較** — 横並びカラム/カード（A vs B、メリット/デメリット）
- **タイムライン / プロセス** — ステップ + 矢印
- **画像ショーケース** — ヒーロー画像、ギャラリー

### フォントサイズ

| 要素 | 推奨 (pt) |
|------|-----------|
| スライドタイトル | 28 〜 36 (`title`) |
| セクションヘッダー | 20 〜 24 |
| 本文テキスト | 14 〜 16 (`body`) |
| キャプション / 出典 | 10 〜 12 (`caption`) |
| データコールアウト | 48 〜 72 |

### 図形を活用したコンテンツ表現

**プログレスバー:**
```javascript
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 3, w: 8, h: 0.25, fill: { color: theme.bg }, rectRadius: 0.125
});
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 3, w: 5.6, h: 0.25, fill: { color: theme.accent }, rectRadius: 0.125
});
```

**KPI カード:**
```javascript
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.5, w: 2.8, h: 2, fill: { color: theme.bg }, rectRadius: 0.1
});
slide.addText("85%", {
  x: 0.5, y: 1.8, w: 2.8, h: 1,
  fontSize: 48, fontFace: "Arial", color: theme.accent, bold: true, align: "center"
});
slide.addText("達成率", {
  x: 0.5, y: 2.7, w: 2.8, h: 0.5,
  fontSize: 14, fontFace: "Noto Sans JP", color: theme.gray, align: "center"
});
```

**アイコン付きリスト:**
```javascript
const items = ["項目その 1", "項目その 2", "項目その 3"];
items.forEach((item, i) => {
  const y = 1.5 + i * 0.8;
  slide.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: theme.accent } });
  slide.addText(String(i + 1), {
    x: 0.8, y, w: 0.5, h: 0.5,
    fontSize: 16, color: theme.white, bold: true, align: "center", valign: "middle"
  });
  slide.addText(item, {
    x: 1.5, y, w: 7, h: 0.5, fontSize: 16, fontFace: "Noto Sans JP", color: theme.dark
  });
});
```

### デザイン原則

1. **本文テキストは左揃え** — 段落やリストをセンタリングしない（センタリングはタイトルのみ）
2. **サイズのコントラスト** — タイトルは 28pt 以上、本文は 14〜16pt
3. **視覚要素は必須** — すべてのコンテンツスライドに最低 1 つの非テキスト要素
4. **余白を確保** — 最小マージン 0.4"、ブロック間 0.3〜0.5"
5. **レイアウトの変化** — 前後のスライドと異なるレイアウトを使う
6. **タイトル下の飾り線は使わない** — AI 生成スライドの典型。余白か背景色で区切る

---

## 5. まとめ / クロージング

- **用途**: まとめ + アクション
- **内容**: 要点の振り返り、次のステップ、感謝

### レイアウト例

**要点振り返り型:**
```
|  まとめ                          |
|  ✓ 要点 1                       |
|  ✓ 要点 2                       |
|  ✓ 要点 3                       |
```

**Thank You 型:** 全面 `theme.primary` + 中央に `theme.white` の大きなテキスト。

### フォントサイズ

| 要素 | 推奨 (pt) |
|------|-----------|
| クロージングタイトル | 44 〜 60 |
| 要点 / アクションアイテム | 18 〜 24 |

---

## 追加レイアウトパターン

視覚的多様性のために活用:

- **2 カラム**: テキスト左、図版右
- **アイコン + テキスト行**: 色付き円 + 太字ヘッダー + 説明
- **2x2 / 2x3 グリッド**: 画像 + コンテンツブロック
- **大きな数値コールアウト**: 60〜72pt の数値 + 小さなラベル
- **比較カラム**: Before/After、メリット/デメリット
- **タイムライン / プロセスフロー**: 番号付きステップ + 矢印
- **KPI カード行**: 3〜4 枚の図形ベースカード
