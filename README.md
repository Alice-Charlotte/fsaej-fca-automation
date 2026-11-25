# FSAEJ FCA Automation

学生フォーミュラ日本大会（FSAEJ）のコスト審査で使う **FCA（Factory Cost Analysis）の計算・集計と、公式テンプレート形式シートの生成を自動化するツール** です。

明細を1行ずつ入力するだけで、公式コストテーブルからの単価引き（パラメータ式の単価計算を含む）、工程倍率の適用、部品別・アセンブリ別の集計まで自動で行い、公式FCAテンプレート形式のシートを部品・アセンブリごとに生成します。

## 特長

- 公式コストテーブル全件対応（Materials / Processes / Fasteners / Process Multipliers / Tooling）
- Size1 / Size2 と係数 C1 / C2 から単価を計算するパラメータ式（材料・締結の価格式パターンを全分類）にも対応
- 材料の断面積計算（Area × Length × Density）や、工具費の PVF / FracIncld 計算など、公式テンプレートの計算式をそのまま再現
- 公式テンプレートのサンプル値で検証済み（Part Cost / Assembly Cost がテンプレート自身の計算と完全一致）

## 必要なもの

このリポジトリには**コードと説明のみ**が含まれます。以下の公式ファイルは主催者（FSAEJ）配布物のため含めていません。各自で入手し、このリポジトリのファイルと同じフォルダに置いてください。

- コストテーブル一式：`tblMaterials_*.xlsx` / `tblProcesses_*.xlsx` / `tblFasteners_*.xlsx` / `tblProcessMultipliers_*.xlsx` / `tblTooling_*.xlsx`
- 公式FCAテンプレート：`FSAEJ_FCA_Template_*.xlsx`

実行環境：Windows / macOS / Linux + Python 3、`openpyxl`（`pip install openpyxl`）

## 使い方

### 1. 入力ブックを作る

```bash
python make_input_workbook.py
```

公式テーブルから `SITFP_FCA_auto_calc_v2.xlsx`（参照テーブル＋入力シート＋集計＋チェック）が生成されます。

### 2. 入力する

生成されたブックの「入力_FCA明細」シートに、黄色い列だけを入力します。

- 項目名は参照シートの英語表記と完全一致させること
- 材料で断面積を使う行は、Area列に数値かExcel式を入れる

単価・倍率・明細コストは自動計算され、「集計_部品別」「集計_アセンブリ」に自動集計されます。

### 3. FCAシートを生成する

```bash
python build_fca_sheets.py
```

Windowsでは `run_FCA.bat` のダブルクリックでも実行できます（初回は openpyxl を自動インストールします）。
同じフォルダに `SITFP_FCA_generated.xlsx` が生成され、部品・アセンブリごとに公式テンプレート形式のシートが入ります。アセンブリ側の部品コストは部品シートへのリンク式になっているため、部品を直せばアセンブリも追従します。

## 制限

- 公式テンプレートの行数上限：部品1つあたり 材料8行・工程8行・締結8行・工具2行、アセンブリの構成部品は6点まで。超えるとエラーで止まります
- 検証は2026年版のテーブル／テンプレートで行っています。新年度版が配布されたら、ファイルを差し替えてください

## ライセンス

MIT License（`LICENSE` を参照）。ただし上記の公式ファイル（コストテーブル・テンプレート）は本リポジトリの対象外で、各配布元の条件に従ってください。
