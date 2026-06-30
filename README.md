# brand-pptx

テンプレート PPTX を複製し、レイアウトを選んでプレースホルダーを差し替えることで、**そのテンプレートのデザインに沿った PowerPoint** を生成する [Agent Skill](https://github.com/vercel-labs/skills) です。

「青系でモダンに」のようなテキスト指示だけでブランドを固定するのは現実的に不可能で、毎回ブレます。
このスキルは配色・装飾・レイアウトを **PPTX テンプレートというデータ** に持たせ、生成コードは中身を流し込むだけにすることで、テンプレートから離れたスライドが出てこないようにします。

- **同梱のクリーンな既定テンプレート**（`assets/template.pptx`, 16:9）ですぐ動きます。
- **自社テンプレート（.pptx）に差し替えれば**、そのブランドをそのまま再現できます。

> もとは社内向けに「自社テンプレート準拠スライドを生成するスキル」を作ったところ好評だったので、ブランド固有の資産を抜いて誰でも使える形にしたものです。

![brand-pptx の既定テンプレートで生成したスライド例（カバー / セクション / KPI カード / エンディング）](docs/preview.png)

## 特長

- **テンプレート準拠** — 配色・装飾・ロゴはテンプレートが持つ。生成はプレースホルダーを埋めるだけ
- **すぐ使える / 自社化も簡単** — 既定テンプレートで即動作。自社テンプレートは Claude に渡すだけで設定
- **視覚パーツの自動配置** — KPI カード・進捗バー・ステップフロー・テーブルを `theme.json` の色で生成
- **編集にも対応** — 既存 PPTX を XML レベルで安全に編集するワークフローを同梱

## インストール

```bash
npx skills add KeMezz/brand-pptx
```

依存パッケージ:

```bash
pip install python-pptx "markitdown[pptx]"
npm install pptxgenjs   # テンプレート外のカスタムスライド用（任意）
```

## 使い方

Claude Code などで、本スキルを呼び出して資料作成を依頼します。

```
四半期サマリーのプレゼン資料を作って
```

初回に **「自社のテンプレート（.pptx）はありますか？」** と一度だけ聞かれます。

- **ある** → その .pptx を渡すだけ。レイアウトを自動で読み取り、`theme.json` を設定して、以後そのブランドで生成します。
- **ない** → 同梱のクリーンなテンプレートでそのまま生成します（設定不要）。

生成した PPTX は、Google スライドで「ファイル」→「開く」→「アップロード」すると Google Workspace でも開けます。

## 自社ブランドへの差し替え（手動で行う場合）

Claude に任せず手動で設定するときは、次の 2 つだけです。

**1. テンプレートを差し込む**

自社の `.pptx` を `assets/` に置き、レイアウト index を確認して `theme.json` に書きます。

```bash
python3 tools/inspect_template.py assets/your-template.pptx
# [ 0] Title Slide / [ 1] Title and Content / [ 2] Section Header ...
```

```jsonc
{
  "template": { "path": "assets/your-template.pptx" },
  "layoutMap": { "cover": 0, "section": 4, "content": 7, "contentVisual": 8, "ending": 1 }
}
```

**2. 視覚パーツの色を合わせる（任意）**

KPI カードや進捗バーなど **コードが描く部分** の色は `theme.json` の `colors` で調整します（`#` を付けない 6 桁 hex）。

```jsonc
{ "colors": { "accent": "5B6CFF", "bg": "F4F5FA", "dark": "131722", "gray": "5A6473" } }
```

> 既定テンプレートは `python3 tools/build_template.py` が `theme.json` の色から生成します。色を変えて再実行すると既定テンプレートの配色も変わります。

**テンプレートに足りない役割がある場合（任意）**: 自社テンプレートにある役割向けのレイアウトが無いときは、`layoutMap` のその役割を `null` にし、`theme.json` の `setup.drawMissingRoles` を `true` にすると、その役割だけ `theme` の色でスライドに直接描画して補います（**テンプレート自体は変更しません**）。既定は `false`。

## 動作確認

```bash
python3 tools/make_sample.py        # examples/sample.pptx を生成
python3 -m markitdown examples/sample.pptx   # 中身をテキストで確認
```

## ライセンス

[MIT](./LICENSE) © 2026 KeMezz

テキストおよびワークフローの一部は [MiniMax-AI/skills の pptx-generator](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator)（MIT License）をベースにしています。詳細は [NOTICE.md](./NOTICE.md) を参照してください。
